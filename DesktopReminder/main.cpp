#define _CRT_SECURE_NO_WARNINGS
#ifndef NOMINMAX
#define NOMINMAX
#endif

#include <windows.h>
#include <windowsx.h>
#include <commctrl.h>
#include <shellapi.h>
#include <shlobj.h>
#include <commdlg.h>
#include <strsafe.h>
#include <uxtheme.h>

#include <algorithm>
#include <cctype>
#include <cmath>
#include <cstdio>
#include <ctime>
#include <deque>
#include <map>
#include <optional>
#include <sstream>
#include <set>
#include <string>
#include <vector>

#include "resource.h"
#include "reminder_core.h"
#include "storage_io.h"

#pragma comment(lib, "comctl32.lib")
#pragma comment(lib, "shell32.lib")
#pragma comment(lib, "comdlg32.lib")
#pragma comment(lib, "uxtheme.lib")
#pragma comment(linker, "/manifestdependency:\"type='win32' name='Microsoft.Windows.Common-Controls' version='6.0.0.0' processorArchitecture='*' publicKeyToken='6595b64144ccf1df' language='*'\"")

namespace {

constexpr UINT WMAPP_TRAY = WM_APP + 1;
constexpr UINT WMAPP_SHOW_NEXT_POPUP = WM_APP + 2;
constexpr UINT WMAPP_AGENT_RELOAD = WM_APP + 3;
constexpr UINT TIMER_REMINDER = 1;
constexpr int MAX_UNDO = 20;
constexpr int DEFAULT_REMINDER_HOUR = 9;
constexpr int REPEAT_NOTIFY_MINUTES = 30;
constexpr UINT MAX_TIMER_WAIT_MS = 24u * 60u * 60u * 1000u;

enum class TaskType { Yearly, Monthly, Daily };
enum class NotifyType { Popup, Tray, Both, None };
enum class TaskColorMode { Auto, White, Yellow, Orange };

enum class RunMode { Agent, Editor };

struct Task {
    int id = 0;
    TaskType type = TaskType::Monthly;
    int month = 1;
    int day = 1;
    int minutesOfDay = DEFAULT_REMINDER_HOUR * 60;
    std::wstring category;
    std::wstring title;
    NotifyType notify = NotifyType::Tray;
    TaskColorMode color = TaskColorMode::Auto;
    bool enabled = true;
    std::time_t lastCompleted = 0;
    std::time_t snoozedUntil = 0;
    std::time_t lastNotified = 0;
    std::wstring cachedDueDisplay;
    std::wstring cachedNotifyDisplay;
};

struct AppState {
    HINSTANCE instance = nullptr;
    std::wstring exePath;
    RunMode mode = RunMode::Editor;
    HWND mainWnd = nullptr;
    HWND tab = nullptr;
    HWND filterCombo = nullptr;
    HWND pendingCheck = nullptr;
    HWND pathLabel = nullptr;
    HWND hintLabel = nullptr;
    HWND emptyLabel = nullptr;
    HWND list = nullptr;
    HWND statusLabel = nullptr;
    HWND settingsTaskLabel = nullptr;
    HWND settingsTaskCombo = nullptr;
    HWND settingsColorLabel = nullptr;
    HWND settingsColorCombo = nullptr;
    HWND settingsApplyButton = nullptr;
    HWND settingsInfoLabel = nullptr;
    HWND settingsRulesLabel = nullptr;
    HFONT font = nullptr;
    HFONT fontBold = nullptr;
    HBRUSH bgBrush = nullptr;
    NOTIFYICONDATAW tray{};
    std::wstring csvPath;
    std::wstring iniPath;
    std::vector<Task> tasks;
    std::vector<std::vector<Task>> undoStack;
    bool onlyPending = false;
    int filterIndex = 0;
    int currentTab = 0;
    bool popupOpen = false;
    std::deque<int> popupQueue;
    bool firstMinimizeNoticeShown = false;
    bool uiCreated = false;
    UINT currentDpi = 96;
    bool exiting = false;
    HANDLE agentMutex = nullptr;
    int csvSkippedRows = 0;
    std::vector<int> csvSkippedLineNumbers;
    std::vector<std::wstring> csvSkippedReasons;
};

struct TaskEditorInitData {
    Task* task = nullptr;
    std::wstring title;
};

AppState g_app;

bool IsAgentMode() { return g_app.mode == RunMode::Agent; }
bool IsEditorMode() { return g_app.mode == RunMode::Editor; }

constexpr const wchar_t* AGENT_WINDOW_CLASS = L"DesktopReminderNativeAgent";
constexpr const wchar_t* EDITOR_WINDOW_CLASS = L"DesktopReminderNativeEditor";

UINT GetSystemDpiSafe() {
    auto getDpiForSystem = reinterpret_cast<UINT(WINAPI*)()>(GetProcAddress(GetModuleHandleW(L"user32.dll"), "GetDpiForSystem"));
    if (getDpiForSystem) {
        UINT dpi = getDpiForSystem();
        if (dpi >= 96) return dpi;
    }
    HDC screen = GetDC(nullptr);
    int dpi = screen ? GetDeviceCaps(screen, LOGPIXELSX) : 96;
    if (screen) ReleaseDC(nullptr, screen);
    return dpi > 0 ? static_cast<UINT>(dpi) : 96U;
}

UINT GetWindowDpiSafe(HWND hwnd) {
    auto getDpiForWindow = reinterpret_cast<UINT(WINAPI*)(HWND)>(GetProcAddress(GetModuleHandleW(L"user32.dll"), "GetDpiForWindow"));
    if (getDpiForWindow && hwnd) {
        UINT dpi = getDpiForWindow(hwnd);
        if (dpi >= 96) return dpi;
    }
    return GetSystemDpiSafe();
}

int ScaleForDpi(int value, UINT dpi) {
    return MulDiv(value, static_cast<int>(dpi), 96);
}

std::wstring Trim(const std::wstring& s) {
    size_t start = 0;
    while (start < s.size() && iswspace(s[start])) ++start;
    size_t end = s.size();
    while (end > start && iswspace(s[end - 1])) --end;
    return s.substr(start, end - start);
}

std::wstring ToLower(std::wstring s) {
    std::transform(s.begin(), s.end(), s.begin(), [](wchar_t c) { return (wchar_t)towlower(c); });
    return s;
}

std::wstring GetExeDir() {
    wchar_t path[MAX_PATH]{};
    GetModuleFileNameW(nullptr, path, MAX_PATH);
    std::wstring p = path;
    size_t pos = p.find_last_of(L"\\/");
    return pos == std::wstring::npos ? L"." : p.substr(0, pos);
}

bool EnsureDirectoryForFile(const std::wstring& path) {
    size_t pos = path.find_last_of(L"\\/");
    if (pos == std::wstring::npos) return true;
    std::wstring dir = path.substr(0, pos);
    if (dir.empty()) return true;
    SHCreateDirectoryExW(nullptr, dir.c_str(), nullptr);
    DWORD attrs = GetFileAttributesW(dir.c_str());
    return attrs != INVALID_FILE_ATTRIBUTES && (attrs & FILE_ATTRIBUTE_DIRECTORY);
}

std::wstring GetDefaultCsvPath() {
    PWSTR docs = nullptr;
    std::wstring result;
    if (SUCCEEDED(SHGetKnownFolderPath(FOLDERID_Documents, KF_FLAG_DEFAULT, nullptr, &docs)) && docs) {
        result = std::wstring(docs) + L"\\tasks.csv";
        CoTaskMemFree(docs);
        return result;
    }
    return GetExeDir() + L"\\tasks.csv";
}

bool IsUnsafeConfigPath(const std::wstring& path) {
    std::wstring value = ToLower(Trim(path));
    if (value.empty()) return true;
    if (value.rfind(L"file:", 0) == 0) return true;              // URL
    if (value.rfind(L"\\\\", 0) == 0) return true;               // UNC
    if (value.rfind(L"\\\\?\\", 0) == 0) return true;            // extended path namespace
    if (value.rfind(L"\\\\?\\unc\\", 0) == 0) return true;       // extended UNC
    if (value.rfind(L"\\\\.\\", 0) == 0) return true;             // device path
    if (value.rfind(L"\\\\?\\globalroot\\", 0) == 0) return true; // global device namespace
    if (value.rfind(L"\\??\\", 0) == 0) return true;             // NT object manager namespace
    if (value.size() >= 2 && value[1] == L':') {
        if (!(value.size() >= 3 && value[2] == L'\\')) return true; // reject drive-relative path (e.g. C:foo.csv)
    }
    return false;
}

std::wstring NormalizeCsvPath(const std::wstring& rawPath) {
    std::wstring trimmed = Trim(rawPath);
    if (trimmed.empty()) return L"";
    wchar_t fullPath[MAX_PATH * 4]{};
    DWORD len = GetFullPathNameW(trimmed.c_str(), ARRAYSIZE(fullPath), fullPath, nullptr);
    if (len == 0 || len >= ARRAYSIZE(fullPath)) return trimmed;
    return fullPath;
}

std::wstring GetSafeCsvPath(const std::wstring& rawPath) {
    std::wstring normalized = NormalizeCsvPath(rawPath);
    if (IsUnsafeConfigPath(normalized)) {
        return GetDefaultCsvPath();
    }
    return normalized;
}


std::wstring GetLegacyIniPath() {
    return GetExeDir() + L"\\DesktopReminder.ini";
}

std::wstring GetDefaultIniPath() {
    PWSTR localAppData = nullptr;
    if (SUCCEEDED(SHGetKnownFolderPath(FOLDERID_LocalAppData, KF_FLAG_DEFAULT, nullptr, &localAppData)) && localAppData) {
        std::wstring path = std::wstring(localAppData) + L"\\DesktopReminder\\DesktopReminder.ini";
        CoTaskMemFree(localAppData);
        return path;
    }
    return GetLegacyIniPath();
}

std::wstring CsvHeader() {
    return storage_io::CsvHeader();
}

std::wstring GetAuditLogPath() {
    return storage_io::BuildAuditLogPath(g_app.iniPath, GetExeDir());
}

void AppendAuditLog(const std::wstring& eventName, const std::wstring& message) {
    storage_io::AppendAuditLog(GetAuditLogPath(), eventName, message);
}

bool WriteTextFileUtf8(const std::wstring& path, const std::wstring& text) {
    return storage_io::WriteTextFileUtf8Atomic(path, text, GetAuditLogPath());
}

bool ReadTextFileUtf8(const std::wstring& path, std::wstring& outText) {
    return storage_io::ReadTextFileUtf8(path, outText);
}

bool EnsureCsvExists(const std::wstring& path) {
    return storage_io::EnsureCsvExists(path, GetAuditLogPath());
}

std::vector<std::wstring> SplitCsvLine(const std::wstring& line) {
    std::vector<std::wstring> result;
    std::wstring current;
    bool inQuotes = false;
    for (size_t i = 0; i < line.size(); ++i) {
        wchar_t c = line[i];
        if (inQuotes) {
            if (c == L'"') {
                if (i + 1 < line.size() && line[i + 1] == L'"') {
                    current.push_back(L'"');
                    ++i;
                } else {
                    inQuotes = false;
                }
            } else {
                current.push_back(c);
            }
        } else {
            if (c == L',') {
                result.push_back(current);
                current.clear();
            } else if (c == L'"') {
                inQuotes = true;
            } else {
                current.push_back(c);
            }
        }
    }
    result.push_back(current);
    return result;
}

std::wstring EscapeCsv(const std::wstring& s) {
    if (s.find_first_of(L",\"\r\n") == std::wstring::npos) return s;
    std::wstring v = L"\"";
    for (wchar_t c : s) {
        if (c == L'"') v += L"\"\"";
        else v.push_back(c);
    }
    v += L"\"";
    return v;
}

std::wstring FormatTimeOfDay(int minutes) {
    if (minutes < 0) return L"";
    int h = minutes / 60;
    int m = minutes % 60;
    wchar_t buf[16]{};
    StringCchPrintfW(buf, 16, L"%02d:%02d", h, m);
    return buf;
}

bool ParseTimeOfDay(const std::wstring& text, int& minutes) {
    std::wstring t = Trim(text);
    if (t.empty()) {
        minutes = DEFAULT_REMINDER_HOUR * 60;
        return true;
    }
    for (auto& ch : t) {
        if (ch == 0xFF1A) ch = L':'; // full-width colon
    }
    int h = -1, m = -1;
    if (swscanf_s(t.c_str(), L"%d:%d", &h, &m) == 2 && h >= 0 && h <= 23 && m >= 0 && m <= 59) {
        minutes = h * 60 + m;
        return true;
    }
    // allow legacy strings like "毎日 09:00", "毎月 15日 09:00", "毎年 1/1 09:00"
    for (size_t i = 0; i + 4 < t.size(); ++i) {
        if (iswdigit(t[i])) {
            std::wstring part = t.substr(i);
            if (swscanf_s(part.c_str(), L"%d:%d", &h, &m) == 2 && h >= 0 && h <= 23 && m >= 0 && m <= 59) {
                minutes = h * 60 + m;
                return true;
            }
        }
    }
    return false;
}

bool ParseInt(const std::wstring& s, int& value) {
    return reminder_core::ParseIntNoThrow(s, value);
}

bool ParseLegacyDueText(const std::wstring& dueText, TaskType type, int& month, int& day, int& minutesOfDay) {
    std::wstring t = Trim(dueText);
    if (t.empty()) return false;
    for (auto& ch : t) {
        if (ch == 0xFF0F) ch = L'/';
        if (ch == 0xFF1A) ch = L':';
    }
    ParseTimeOfDay(t, minutesOfDay);
    int a = 0, b = 0;
    if (type == TaskType::Yearly) {
        if (swscanf_s(t.c_str(), L"毎年 %d/%d", &a, &b) == 2 || swscanf_s(t.c_str(), L"%d/%d", &a, &b) == 2) {
            month = a; day = b; return true;
        }
        if (swscanf_s(t.c_str(), L"毎年 %d月%d日", &a, &b) == 2 || swscanf_s(t.c_str(), L"%d月%d日", &a, &b) == 2) {
            month = a; day = b; return true;
        }
    } else if (type == TaskType::Monthly) {
        if (swscanf_s(t.c_str(), L"毎月 %d日", &a) == 1 || swscanf_s(t.c_str(), L"%d日", &a) == 1) {
            day = a; return true;
        }
    } else if (type == TaskType::Daily) {
        return true; // time already parsed
    }
    return false;
}


int DaysInMonth(int year, int month) {
    return reminder_core::DaysInMonth(year, month);
}


int ClampInt(int value, int minimum, int maximum) {
    return reminder_core::ClampInt(value, minimum, maximum);
}

void NormalizeTask(Task& task) {
    if (task.id <= 0) task.id = 1;
    if (task.minutesOfDay < 0 || task.minutesOfDay > 23 * 60 + 59) task.minutesOfDay = DEFAULT_REMINDER_HOUR * 60;
    switch (task.type) {
    case TaskType::Yearly:
        task.month = ClampInt(task.month, 1, 12);
        task.day = ClampInt(task.day, 1, DaysInMonth(2024, task.month));
        break;
    case TaskType::Monthly:
        task.month = 1;
        task.day = ClampInt(task.day, 1, 31);
        break;
    case TaskType::Daily:
    default:
        task.month = 1;
        task.day = 1;
        break;
    }
}

std::optional<std::time_t> MakeLocalTime(int year, int month, int day, int hour, int minute, int second = 0) {
    return reminder_core::MakeLocalTime(year, month, day, hour, minute, second);
}

std::tm LocalTm(std::time_t t) {
    std::tm tm{};
    localtime_s(&tm, &t);
    return tm;
}

std::wstring FormatDateTime(std::time_t t) {
    if (t <= 0) return L"";
    auto tm = LocalTm(t);
    wchar_t buf[32]{};
    StringCchPrintfW(buf, 32, L"%04d-%02d-%02d %02d:%02d:%02d",
        tm.tm_year + 1900, tm.tm_mon + 1, tm.tm_mday, tm.tm_hour, tm.tm_min, tm.tm_sec);
    return buf;
}

std::time_t ParseDateTime(const std::wstring& s) {
    std::wstring t = Trim(s);
    if (t.empty()) return 0;
    int y = 0, mo = 0, d = 0, h = 0, mi = 0, se = 0;
    if (swscanf_s(t.c_str(), L"%d-%d-%d %d:%d:%d", &y, &mo, &d, &h, &mi, &se) == 6) {
        auto v = MakeLocalTime(y, mo, d, h, mi, se);
        return v.value_or(0);
    }
    return 0;
}

std::wstring TaskTypeToCsv(TaskType t) {
    switch (t) {
    case TaskType::Yearly: return L"yearly";
    case TaskType::Monthly: return L"monthly";
    default: return L"daily";
    }
}

TaskType CsvToTaskType(const std::wstring& s) {
    std::wstring v = ToLower(Trim(s));
    if (v == L"yearly" || v == L"year" || v == L"年次") return TaskType::Yearly;
    if (v == L"daily" || v == L"day" || v == L"日次") return TaskType::Daily;
    if (v == L"monthly" || v == L"month" || v == L"月次") return TaskType::Monthly;
    return TaskType::Monthly;
}

bool TryCsvToTaskType(const std::wstring& s, TaskType& outType) {
    std::wstring v = ToLower(Trim(s));
    if (v.empty()) return false;
    if (v == L"yearly" || v == L"year" || v == L"年次") {
        outType = TaskType::Yearly;
        return true;
    }
    if (v == L"daily" || v == L"day" || v == L"日次") {
        outType = TaskType::Daily;
        return true;
    }
    if (v == L"monthly" || v == L"month" || v == L"月次") {
        outType = TaskType::Monthly;
        return true;
    }
    return false;
}

std::wstring TaskTypeToDisplay(TaskType t) {
    switch (t) {
    case TaskType::Yearly: return L"年次";
    case TaskType::Monthly: return L"月次";
    default: return L"日次";
    }
}

std::wstring NotifyTypeToCsv(NotifyType n) {
    switch (n) {
    case NotifyType::Popup: return L"popup";
    case NotifyType::Tray: return L"tray";
    case NotifyType::Both: return L"both";
    default: return L"none";
    }
}

NotifyType CsvToNotifyType(const std::wstring& s) {
    std::wstring v = ToLower(Trim(s));
    if (v == L"popup" || v == L"ポップアップ") return NotifyType::Popup;
    if (v == L"both" || v == L"両方") return NotifyType::Both;
    if (v == L"none" || v == L"なし") return NotifyType::None;
    if (v == L"tray" || v == L"通知領域") return NotifyType::Tray;
    return NotifyType::Tray;
}

std::wstring NotifyTypeToDisplay(NotifyType n) {
    switch (n) {
    case NotifyType::Popup: return L"ポップアップ";
    case NotifyType::Tray: return L"通知領域";
    case NotifyType::Both: return L"両方";
    default: return L"なし";
    }
}

std::wstring TaskColorModeToCsv(TaskColorMode c) {
    switch (c) {
    case TaskColorMode::White: return L"white";
    case TaskColorMode::Yellow: return L"yellow";
    case TaskColorMode::Orange: return L"orange";
    default: return L"auto";
    }
}

TaskColorMode CsvToTaskColorMode(const std::wstring& s) {
    std::wstring v = ToLower(Trim(s));
    if (v == L"white" || v == L"白") return TaskColorMode::White;
    if (v == L"yellow" || v == L"黄色") return TaskColorMode::Yellow;
    if (v == L"orange" || v == L"オレンジ") return TaskColorMode::Orange;
    return TaskColorMode::Auto;
}

std::wstring TaskColorModeToDisplay(TaskColorMode c) {
    switch (c) {
    case TaskColorMode::White: return L"白";
    case TaskColorMode::Yellow: return L"黄色";
    case TaskColorMode::Orange: return L"オレンジ";
    default: return L"自動";
    }
}

std::wstring BoolText(bool v) {
    return v ? L"有効" : L"無効";
}

std::wstring BuildTaskDisplayName(const Task& task) {
    std::wstring text = L"[" + TaskTypeToDisplay(task.type) + L"] ";
    if (!Trim(task.category).empty()) text += task.category + L" / ";
    text += task.title;
    return text;
}

void ShowControl(HWND hwnd, bool visible) {
    if (hwnd) ShowWindow(hwnd, visible ? SW_SHOW : SW_HIDE);
}

COLORREF BlendColor(COLORREF base, COLORREF overlay, BYTE overlayAlpha) {
    BYTE br = GetRValue(base), bg = GetGValue(base), bb = GetBValue(base);
    BYTE orr = GetRValue(overlay), og = GetGValue(overlay), ob = GetBValue(overlay);
    BYTE rr = static_cast<BYTE>((br * (255 - overlayAlpha) + orr * overlayAlpha) / 255);
    BYTE rg = static_cast<BYTE>((bg * (255 - overlayAlpha) + og * overlayAlpha) / 255);
    BYTE rb = static_cast<BYTE>((bb * (255 - overlayAlpha) + ob * overlayAlpha) / 255);
    return RGB(rr, rg, rb);
}

bool ParseEnabledText(const std::wstring& s, bool defaultValue = true) {
    std::wstring v = ToLower(Trim(s));
    if (v.empty()) return defaultValue;
    if (v == L"false" || v == L"0" || v == L"無効") return false;
    if (v == L"true" || v == L"1" || v == L"有効") return true;
    return defaultValue;
}

std::wstring NormalizeHeaderName(std::wstring s) {
    s = ToLower(Trim(s));
    std::wstring out;
    out.reserve(s.size());
    for (wchar_t ch : s) {
        if (iswspace(ch)) continue;
        if (ch == 0x3000) continue;
        out.push_back(ch);
    }
    return out;
}

bool GetCurrentPeriod(const Task& task, std::time_t now, std::time_t& start, std::time_t& end) {
    auto tm = LocalTm(now);
    switch (task.type) {
    case TaskType::Daily: {
        auto s = MakeLocalTime(tm.tm_year + 1900, tm.tm_mon + 1, tm.tm_mday, 0, 0, 0);
        auto e = MakeLocalTime(tm.tm_year + 1900, tm.tm_mon + 1, tm.tm_mday + 1, 0, 0, 0);
        if (!s || !e) return false;
        start = *s; end = *e; return true;
    }
    case TaskType::Monthly: {
        auto s = MakeLocalTime(tm.tm_year + 1900, tm.tm_mon + 1, 1, 0, 0, 0);
        int nextYear = tm.tm_year + 1900;
        int nextMonth = tm.tm_mon + 2;
        if (nextMonth == 13) { nextMonth = 1; ++nextYear; }
        auto e = MakeLocalTime(nextYear, nextMonth, 1, 0, 0, 0);
        if (!s || !e) return false;
        start = *s; end = *e; return true;
    }
    case TaskType::Yearly: {
        int year = tm.tm_year + 1900;
        auto s = MakeLocalTime(year, 1, 1, 0, 0, 0);
        auto e = MakeLocalTime(year + 1, 1, 1, 0, 0, 0);
        if (!s || !e) return false;
        start = *s; end = *e; return true;
    }
    }
    return false;
}

bool IsCompletedForCurrentPeriod(const Task& task, std::time_t now) {
    if (task.lastCompleted <= 0) return false;
    std::time_t start = 0, end = 0;
    if (!GetCurrentPeriod(task, now, start, end)) return false;
    return task.lastCompleted >= start && task.lastCompleted < end;
}

struct DueSpec {
    int year = 0;
    int month = 1;
    int day = 1;
    int hour = 0;
    int minute = 0;
};

std::optional<DueSpec> ComputeDueSpecForCurrentPeriod(TaskType type, int monthValue, int dayValue, int minutesOfDay, const std::tm& nowTm) {
    reminder_core::ScheduleType scheduleType = reminder_core::ScheduleType::Monthly;
    if (type == TaskType::Daily) scheduleType = reminder_core::ScheduleType::Daily;
    else if (type == TaskType::Yearly) scheduleType = reminder_core::ScheduleType::Yearly;

    auto coreSpec = reminder_core::ComputeDueSpecForCurrentPeriod(scheduleType, monthValue, dayValue, minutesOfDay, nowTm);
    if (!coreSpec) return std::nullopt;

    DueSpec spec{};
    spec.year = coreSpec->year;
    spec.month = coreSpec->month;
    spec.day = coreSpec->day;
    spec.hour = coreSpec->hour;
    spec.minute = coreSpec->minute;
    return spec;
}

std::optional<std::time_t> GetDueTimeForCurrentPeriod(const Task& task, std::time_t now) {
    auto tm = LocalTm(now);
    auto dueSpec = ComputeDueSpecForCurrentPeriod(task.type, task.month, task.day, task.minutesOfDay, tm);
    if (!dueSpec) return std::nullopt;
    return MakeLocalTime(dueSpec->year, dueSpec->month, dueSpec->day, dueSpec->hour, dueSpec->minute, 0);
}

std::wstring DueDisplay(const Task& sourceTask) {
    auto now = std::time(nullptr);
    auto due = GetDueTimeForCurrentPeriod(sourceTask, now);
    Task task = sourceTask;
    NormalizeTask(task);
    if (due) {
        auto tm = LocalTm(*due);
        task.month = tm.tm_mon + 1;
        task.day = tm.tm_mday;
        task.minutesOfDay = tm.tm_hour * 60 + tm.tm_min;
    }
    std::wstring timeText = FormatTimeOfDay(task.minutesOfDay);
    if (timeText.empty()) timeText = L"09:00";
    switch (task.type) {
    case TaskType::Yearly:
        return L"毎年 " + std::to_wstring(task.month) + L"/" + std::to_wstring(task.day) + L" " + timeText;
    case TaskType::Monthly:
        return L"毎月 " + std::to_wstring(task.day) + L"日 " + timeText;
    default:
        return L"毎日 " + timeText;
    }
}

void RebuildTaskCache(Task& task) {
    task.cachedDueDisplay = DueDisplay(task);
    task.cachedNotifyDisplay = NotifyTypeToDisplay(task.notify);
}

void RebuildAllTaskCache() {
    for (auto& task : g_app.tasks) RebuildTaskCache(task);
}

int NextTaskId() {
    int maxId = 0;
    for (const auto& t : g_app.tasks) maxId = std::max(maxId, t.id);
    return maxId + 1;
}

Task* FindTaskById(int id) {
    for (auto& t : g_app.tasks) if (t.id == id) return &t;
    return nullptr;
}

const Task* FindTaskByIdConst(int id) {
    for (const auto& t : g_app.tasks) if (t.id == id) return &t;
    return nullptr;
}

void PushUndo() {
    g_app.undoStack.push_back(g_app.tasks);
    if ((int)g_app.undoStack.size() > MAX_UNDO) {
        g_app.undoStack.erase(g_app.undoStack.begin());
    }
}

bool SaveTasks();
void RefreshList(int preferTaskId = 0);
void RebuildTaskCache(Task& task);
void RebuildAllTaskCache();
bool LaunchEditorProcess();
HWND FindAgentWindow();
void NotifyAgentReload();
bool EnsureAgentProcess();
void UpdateStatus();
void CheckReminders();
void QueuePopupIfNeeded(int taskId);
void ShowNextPopup();
void ScheduleReminderTimer(bool runImmediateCheck = false);
void DestroyMainControls();
bool EnsureMainControls();
void CreateMainControls(HWND hwnd);
void LayoutMainControls(HWND hwnd);
void PopulateSettingsTaskCombo(int preferTaskId = 0);
void SyncSettingsPanelFromTask(int taskId);
void UpdateVisibleTab();
void SetCurrentTab(int tabIndex);
int SelectedTaskId();

void SaveConfig() {
    WritePrivateProfileStringW(L"General", L"CsvPath", g_app.csvPath.c_str(), g_app.iniPath.c_str());
    WritePrivateProfileStringW(L"General", L"FilterIndex", std::to_wstring(g_app.filterIndex).c_str(), g_app.iniPath.c_str());
    WritePrivateProfileStringW(L"General", L"OnlyPending", g_app.onlyPending ? L"1" : L"0", g_app.iniPath.c_str());
    WritePrivateProfileStringW(L"General", L"CurrentTab", std::to_wstring(g_app.currentTab).c_str(), g_app.iniPath.c_str());
}

void LoadConfig() {
    wchar_t buf[MAX_PATH * 2]{};
    GetPrivateProfileStringW(L"General", L"CsvPath", GetDefaultCsvPath().c_str(), buf, ARRAYSIZE(buf), g_app.iniPath.c_str());
    g_app.csvPath = GetSafeCsvPath(buf);
    g_app.filterIndex = GetPrivateProfileIntW(L"General", L"FilterIndex", 0, g_app.iniPath.c_str());
    g_app.onlyPending = GetPrivateProfileIntW(L"General", L"OnlyPending", 0, g_app.iniPath.c_str()) != 0;
    g_app.currentTab = ClampInt(GetPrivateProfileIntW(L"General", L"CurrentTab", 0, g_app.iniPath.c_str()), 0, 1);
}

bool LoadTasks() {
    g_app.csvSkippedRows = 0;
    g_app.csvSkippedLineNumbers.clear();
    g_app.csvSkippedReasons.clear();
    if (!EnsureCsvExists(g_app.csvPath)) return false;
    std::wstring text;
    if (!ReadTextFileUtf8(g_app.csvPath, text)) return false;
    std::wstringstream ss(text);
    std::wstring line;
    std::vector<Task> loaded;
    std::set<int> usedIds;
    int nextId = 1;
    auto reserveId = [&]() {
        while (usedIds.count(nextId) != 0) ++nextId;
        int id = nextId++;
        usedIds.insert(id);
        return id;
    };
    auto adoptId = [&](int requestedId) {
        if (requestedId > 0 && usedIds.count(requestedId) == 0) {
            usedIds.insert(requestedId);
            if (requestedId >= nextId) nextId = requestedId + 1;
            return requestedId;
        }
        return reserveId();
    };

    struct ColumnMap {
        int id = -1;
        int type = -1;
        int month = -1;
        int day = -1;
        int time = -1;
        int due = -1;
        int category = -1;
        int title = -1;
        int notify = -1;
        int color = -1;
        int enabled = -1;
        int status = -1;
        int lastCompleted = -1;
        int snoozedUntil = -1;
        int lastNotified = -1;
        bool hasHeader = false;
    } map;

    auto assignHeader = [&](const std::vector<std::wstring>& cols) {
        for (int i = 0; i < (int)cols.size(); ++i) {
            std::wstring h = NormalizeHeaderName(cols[i]);
            if (h == L"id") map.id = i;
            else if (h == L"type" || h == L"単位") map.type = i;
            else if (h == L"month" || h == L"月") map.month = i;
            else if (h == L"day" || h == L"日") map.day = i;
            else if (h == L"time" || h == L"時刻" || h == L"時間") map.time = i;
            else if (h == L"due" || h == L"予定") map.due = i;
            else if (h == L"category" || h == L"カテゴリ") map.category = i;
            else if (h == L"title" || h == L"task" || h == L"タスク") map.title = i;
            else if (h == L"notify" || h == L"通知") map.notify = i;
            else if (h == L"color" || h == L"色") map.color = i;
            else if (h == L"enabled" || h == L"有効") map.enabled = i;
            else if (h == L"status" || h == L"状態") map.status = i;
            else if (h == L"last_completed") map.lastCompleted = i;
            else if (h == L"snoozed_until" || h == L"スヌーズ") map.snoozedUntil = i;
            else if (h == L"last_notified") map.lastNotified = i;
        }
        map.hasHeader = (map.type >= 0 || map.title >= 0 || map.due >= 0 || map.time >= 0 || map.id >= 0);
    };

    std::vector<std::wstring> bufferedFirstRow;
    bool firstDataBuffered = false;
    bool firstLineHandled = false;
    int lineNumber = 0;
    int bufferedFirstRowLine = 0;

    while (std::getline(ss, line)) {
        ++lineNumber;
        if (!line.empty() && line.back() == L'\r') line.pop_back();
        if (Trim(line).empty()) continue;
        auto cols = SplitCsvLine(line);
        if (!firstLineHandled) {
            firstLineHandled = true;
            assignHeader(cols);
            if (!map.hasHeader) {
                bufferedFirstRow = cols;
                firstDataBuffered = true;
                bufferedFirstRowLine = lineNumber;
            }
            continue;
        }

        auto markSkipped = [&](int rowLineNumber, const std::wstring& reason) {
            ++g_app.csvSkippedRows;
            g_app.csvSkippedLineNumbers.push_back(rowLineNumber);
            g_app.csvSkippedReasons.push_back(L"行" + std::to_wstring(rowLineNumber) + L":" + reason);
            AppendAuditLog(L"CSV_ROW_SKIPPED",
                L"line=" + std::to_wstring(rowLineNumber) + L", reason=" + reason + L", path=" + g_app.csvPath);
        };

        auto processCols = [&](const std::vector<std::wstring>& row, int rowLineNumber) {
            if (row.empty()) return;
            Task task;
            task.id = 0;
            task.type = TaskType::Monthly;
            task.month = 1;
            task.day = 1;
            task.minutesOfDay = DEFAULT_REMINDER_HOUR * 60;
            task.notify = NotifyType::Tray;
            task.color = TaskColorMode::Auto;
            task.enabled = true;

            auto get = [&](int idx) -> std::wstring {
                return (idx >= 0 && idx < (int)row.size()) ? row[idx] : L"";
            };

            if (map.hasHeader) {
                std::wstring idText = get(map.id);
                if (!Trim(idText).empty() && !ParseInt(idText, task.id)) {
                    markSkipped(rowLineNumber, L"invalid id");
                    return;
                }

                std::wstring typeText = get(map.type);
                if (!Trim(typeText).empty()) {
                    TaskType parsedType;
                    if (!TryCsvToTaskType(typeText, parsedType)) {
                        markSkipped(rowLineNumber, L"invalid type");
                        return;
                    }
                    task.type = parsedType;
                }
                std::wstring monthText = get(map.month);
                std::wstring dayText = get(map.day);
                if (!Trim(monthText).empty() && !ParseInt(monthText, task.month)) {
                    markSkipped(rowLineNumber, L"invalid month");
                    return;
                }
                if (!Trim(dayText).empty() && !ParseInt(dayText, task.day)) {
                    markSkipped(rowLineNumber, L"invalid day");
                    return;
                }

                std::wstring dueText = get(map.due);
                std::wstring timeText = get(map.time);
                int mod = DEFAULT_REMINDER_HOUR * 60;
                bool parsedTime = ParseTimeOfDay(timeText, mod);
                if (!parsedTime && !Trim(dueText).empty()) parsedTime = ParseTimeOfDay(dueText, mod);
                if (!parsedTime && !Trim(timeText).empty()) {
                    markSkipped(rowLineNumber, L"invalid time");
                    return;
                }
                task.minutesOfDay = mod;
                if (!Trim(dueText).empty()) {
                    ParseLegacyDueText(dueText, task.type, task.month, task.day, task.minutesOfDay);
                }

                task.category = get(map.category);
                task.title = get(map.title);
                task.notify = CsvToNotifyType(get(map.notify));
                task.color = CsvToTaskColorMode(get(map.color));
                task.enabled = ParseEnabledText(get(map.enabled), true);
                task.lastCompleted = ParseDateTime(get(map.lastCompleted));
                task.snoozedUntil = ParseDateTime(get(map.snoozedUntil));
                task.lastNotified = ParseDateTime(get(map.lastNotified));

                if (Trim(task.title).empty()) {
                    markSkipped(rowLineNumber, L"empty title");
                    return;
                }
            } else {
                if (row.size() < 2) {
                    markSkipped(rowLineNumber, L"too few columns");
                    return;
                }
                if (!row[0].empty() && !ParseInt(row[0], task.id)) {
                    markSkipped(rowLineNumber, L"invalid id");
                    return;
                }
                task.type = CsvToTaskType(row.size() > 1 ? row[1] : L"");
                if (row.size() > 2 && !Trim(row[2]).empty() && !ParseInt(row[2], task.month)) {
                    markSkipped(rowLineNumber, L"invalid month");
                    return;
                }
                if (row.size() > 3 && !Trim(row[3]).empty() && !ParseInt(row[3], task.day)) {
                    markSkipped(rowLineNumber, L"invalid day");
                    return;
                }
                int mod = DEFAULT_REMINDER_HOUR * 60;
                std::wstring timeCol = row.size() > 4 ? row[4] : L"";
                bool parsedTime = ParseTimeOfDay(timeCol, mod);
                if (!parsedTime && !Trim(timeCol).empty()) {
                    markSkipped(rowLineNumber, L"invalid time");
                    return;
                }
                task.minutesOfDay = mod;
                if ((!parsedTime || Trim(timeCol).empty()) && row.size() > 1) {
                    ParseLegacyDueText(row[1], task.type, task.month, task.day, task.minutesOfDay);
                }
                task.category = row.size() > 5 ? row[5] : L"";
                task.title = row.size() > 6 ? row[6] : L"";
                task.notify = row.size() > 7 ? CsvToNotifyType(row[7]) : NotifyType::Tray;
                bool hasColorColumn = row.size() >= 13;
                task.color = hasColorColumn && row.size() > 8 ? CsvToTaskColorMode(row[8]) : TaskColorMode::Auto;
                task.enabled = hasColorColumn ? ParseEnabledText(row.size() > 9 ? row[9] : L"", true)
                                            : !(row.size() > 8 && ToLower(Trim(row[8])) == L"false");
                task.lastCompleted = hasColorColumn ? (row.size() > 10 ? ParseDateTime(row[10]) : 0)
                                                    : (row.size() > 9 ? ParseDateTime(row[9]) : 0);
                task.snoozedUntil = hasColorColumn ? (row.size() > 11 ? ParseDateTime(row[11]) : 0)
                                                   : (row.size() > 10 ? ParseDateTime(row[10]) : 0);
                task.lastNotified = hasColorColumn ? (row.size() > 12 ? ParseDateTime(row[12]) : 0)
                                                   : (row.size() > 11 ? ParseDateTime(row[11]) : 0);
                if (Trim(task.title).empty()) {
                    markSkipped(rowLineNumber, L"empty title");
                    return;
                }
            }

            task.id = adoptId(task.id);
            NormalizeTask(task);
            loaded.push_back(task);
        };

        if (firstDataBuffered) {
            processCols(bufferedFirstRow, bufferedFirstRowLine);
            firstDataBuffered = false;
        }
        processCols(cols, lineNumber);
    }

    if (firstDataBuffered) {
        Task task;
        // single-line CSV without header/data separation is unusual, but still attempt to load it.
        map.hasHeader = false;
        auto cols = bufferedFirstRow;
        auto skipBufferedRow = [&](const std::wstring& reason) {
            AppendAuditLog(L"CSV_ROW_SKIPPED",
                L"line=" + std::to_wstring(bufferedFirstRowLine) + L", reason=" + reason + L", path=" + g_app.csvPath);
            ++g_app.csvSkippedRows;
            g_app.csvSkippedLineNumbers.push_back(bufferedFirstRowLine);
            g_app.csvSkippedReasons.push_back(L"行" + std::to_wstring(bufferedFirstRowLine) + L":" + reason);
            g_app.tasks = std::move(loaded);
            RebuildAllTaskCache();
            if (g_app.mainWnd && IsAgentMode()) ScheduleReminderTimer(false);
            return true;
        };
        if (!cols.empty()) {
            if (cols.size() < 2) {
                return skipBufferedRow(L"too few columns");
            }
            if (!cols[0].empty() && !ParseInt(cols[0], task.id)) {
                return skipBufferedRow(L"invalid id");
            }
            task.type = CsvToTaskType(cols.size() > 1 ? cols[1] : L"");
            if (cols.size() > 2 && !Trim(cols[2]).empty() && !ParseInt(cols[2], task.month)) {
                return skipBufferedRow(L"invalid month");
            }
            if (cols.size() > 3 && !Trim(cols[3]).empty() && !ParseInt(cols[3], task.day)) {
                return skipBufferedRow(L"invalid day");
            }
            int mod = DEFAULT_REMINDER_HOUR * 60;
            std::wstring timeCol = cols.size() > 4 ? cols[4] : L"";
            if (!ParseTimeOfDay(timeCol, mod) && !Trim(timeCol).empty()) {
                return skipBufferedRow(L"invalid time");
            }
            task.minutesOfDay = mod;
            if (cols.size() > 1) ParseLegacyDueText(cols[1], task.type, task.month, task.day, task.minutesOfDay);
            task.category = cols.size() > 5 ? cols[5] : L"";
            task.title = cols.size() > 6 ? cols[6] : L"";
            task.notify = cols.size() > 7 ? CsvToNotifyType(cols[7]) : NotifyType::Tray;
            bool hasColorColumn = cols.size() >= 13;
            task.color = hasColorColumn && cols.size() > 8 ? CsvToTaskColorMode(cols[8]) : TaskColorMode::Auto;
            task.enabled = hasColorColumn ? ParseEnabledText(cols.size() > 9 ? cols[9] : L"", true)
                                          : !(cols.size() > 8 && ToLower(Trim(cols[8])) == L"false");
            task.lastCompleted = hasColorColumn ? (cols.size() > 10 ? ParseDateTime(cols[10]) : 0)
                                                : (cols.size() > 9 ? ParseDateTime(cols[9]) : 0);
            task.snoozedUntil = hasColorColumn ? (cols.size() > 11 ? ParseDateTime(cols[11]) : 0)
                                               : (cols.size() > 10 ? ParseDateTime(cols[10]) : 0);
            task.lastNotified = hasColorColumn ? (cols.size() > 12 ? ParseDateTime(cols[12]) : 0)
                                               : (cols.size() > 11 ? ParseDateTime(cols[11]) : 0);
            if (Trim(task.title).empty()) {
                return skipBufferedRow(L"empty title");
            }
            task.id = adoptId(task.id);
            NormalizeTask(task);
            loaded.push_back(task);
        }
    }

    g_app.tasks = std::move(loaded);
    RebuildAllTaskCache();
    if (g_app.mainWnd && IsAgentMode()) ScheduleReminderTimer(false);
    return true;
}

bool SaveTasks() {
    std::wstring text = CsvHeader();
    for (const auto& originalTask : g_app.tasks) {
        Task task = originalTask;
        NormalizeTask(task);
        std::wstring line;
        line += std::to_wstring(task.id) + L",";
        line += EscapeCsv(TaskTypeToCsv(task.type)) + L",";
        line += (task.type == TaskType::Daily ? L"" : std::to_wstring(task.month)) + L",";
        line += (task.type == TaskType::Daily ? L"" : std::to_wstring(task.day)) + L",";
        line += EscapeCsv(FormatTimeOfDay(task.minutesOfDay)) + L",";
        line += EscapeCsv(task.category) + L",";
        line += EscapeCsv(task.title) + L",";
        line += EscapeCsv(NotifyTypeToCsv(task.notify)) + L",";
        line += EscapeCsv(TaskColorModeToCsv(task.color)) + L",";
        line += task.enabled ? L"true" : L"false";
        line += L"," + EscapeCsv(FormatDateTime(task.lastCompleted));
        line += L"," + EscapeCsv(FormatDateTime(task.snoozedUntil));
        line += L"," + EscapeCsv(FormatDateTime(task.lastNotified));
        line += L"\n";
        text += line;
    }
    bool ok = WriteTextFileUtf8(g_app.csvPath, text);
    if (!ok) {
        AppendAuditLog(L"SAVE_FAILED", L"Failed to save CSV: " + g_app.csvPath);
    }
    if (ok) {
        if (g_app.mainWnd && IsAgentMode()) ScheduleReminderTimer(false);
        if (IsEditorMode()) NotifyAgentReload();
    }
    return ok;
}

void SetPathLabel() {
    if (!g_app.pathLabel) return;
    std::wstring text = L"保存先: " + g_app.csvPath;
    SetWindowTextW(g_app.pathLabel, text.c_str());
}

void UpdateStatus() {
    if (!g_app.statusLabel) return;
    int total = (int)g_app.tasks.size();
    int visible = g_app.list ? ListView_GetItemCount(g_app.list) : 0;
    int pending = 0;
    auto now = std::time(nullptr);
    for (const auto& t : g_app.tasks) if (!IsCompletedForCurrentPeriod(t, now)) ++pending;
    wchar_t buf[160]{};
    StringCchPrintfW(buf, 160, L"表示 %d 件 / 全体 %d 件 / 未完了 %d 件", visible, total, pending);
    std::wstring status = buf;
    if (g_app.csvSkippedRows > 0) {
        status += L" / 読込スキップ " + std::to_wstring(g_app.csvSkippedRows) + L" 件";
        if (!g_app.csvSkippedLineNumbers.empty()) {
            status += L" (行:";
            const size_t maxDisplay = 3;
            for (size_t i = 0; i < g_app.csvSkippedLineNumbers.size() && i < maxDisplay; ++i) {
                if (i > 0) status += L",";
                status += std::to_wstring(g_app.csvSkippedLineNumbers[i]);
            }
            if (g_app.csvSkippedLineNumbers.size() > maxDisplay) status += L",...";
            status += L")";
        }
        if (!g_app.csvSkippedReasons.empty()) {
            status += L" [" + g_app.csvSkippedReasons.front() + L"]";
        }
    }
    SetWindowTextW(g_app.statusLabel, status.c_str());
    if (g_app.emptyLabel) {
        ShowWindow(g_app.emptyLabel, visible == 0 ? SW_SHOW : SW_HIDE);
    }
}

COLORREF ColorFromTaskColorMode(TaskColorMode mode) {
    switch (mode) {
    case TaskColorMode::Yellow: return RGB(255, 248, 196);
    case TaskColorMode::Orange: return RGB(255, 226, 163);
    case TaskColorMode::White:
    case TaskColorMode::Auto:
    default:
        return RGB(255, 255, 255);
    }
}

COLORREF ApplyRowStripe(COLORREF baseColor, int rowIndex) {
    if ((rowIndex % 2) == 0) return baseColor;
    return BlendColor(baseColor, RGB(232, 245, 252), 96);
}

TaskColorMode ResolveAutomaticTaskColor(const Task& task, std::time_t now) {
    if (IsCompletedForCurrentPeriod(task, now)) return TaskColorMode::White;
    auto due = GetDueTimeForCurrentPeriod(task, now);
    if (!due) return TaskColorMode::White;

    double diffSeconds = std::difftime(now, *due);
    double absDiffSeconds = std::fabs(diffSeconds);

    if (task.type == TaskType::Daily) {
        const double warnSeconds = 10.0 * 60.0;
        if (diffSeconds > warnSeconds) return TaskColorMode::Orange;
        if (diffSeconds >= 0.0 || absDiffSeconds <= warnSeconds) return TaskColorMode::Yellow;
        return TaskColorMode::White;
    }

    const double warnSeconds = 24.0 * 60.0 * 60.0;
    if (diffSeconds > warnSeconds) return TaskColorMode::Orange;
    if (diffSeconds >= 0.0 || absDiffSeconds <= warnSeconds) return TaskColorMode::Yellow;
    return TaskColorMode::White;
}

TaskColorMode ResolveEffectiveTaskColorMode(const Task& task, std::time_t now) {
    if (task.color != TaskColorMode::Auto) return task.color;
    return ResolveAutomaticTaskColor(task, now);
}

COLORREF GetTaskRowBackColor(const Task& task, std::time_t now, int rowIndex) {
    return ApplyRowStripe(ColorFromTaskColorMode(ResolveEffectiveTaskColorMode(task, now)), rowIndex);
}

void PopulateSettingsTaskCombo(int preferTaskId) {
    if (!g_app.settingsTaskCombo) return;
    int keepTaskId = preferTaskId;
    if (keepTaskId == 0) {
        int sel = (int)SendMessageW(g_app.settingsTaskCombo, CB_GETCURSEL, 0, 0);
        if (sel >= 0) keepTaskId = (int)SendMessageW(g_app.settingsTaskCombo, CB_GETITEMDATA, sel, 0);
    }
    SendMessageW(g_app.settingsTaskCombo, CB_RESETCONTENT, 0, 0);
    int selectedIndex = -1;
    for (const auto& task : g_app.tasks) {
        std::wstring label = BuildTaskDisplayName(task);
        int idx = (int)SendMessageW(g_app.settingsTaskCombo, CB_ADDSTRING, 0, (LPARAM)label.c_str());
        SendMessageW(g_app.settingsTaskCombo, CB_SETITEMDATA, idx, task.id);
        if (keepTaskId != 0 && task.id == keepTaskId) selectedIndex = idx;
    }
    if (selectedIndex < 0 && !g_app.tasks.empty()) selectedIndex = 0;
    SendMessageW(g_app.settingsTaskCombo, CB_SETCURSEL, selectedIndex, 0);
}

void SyncSettingsPanelFromTask(int taskId) {
    if (!g_app.settingsTaskCombo || !g_app.settingsColorCombo || !g_app.settingsInfoLabel) return;

    int targetIndex = -1;
    int count = (int)SendMessageW(g_app.settingsTaskCombo, CB_GETCOUNT, 0, 0);
    for (int i = 0; i < count; ++i) {
        if ((int)SendMessageW(g_app.settingsTaskCombo, CB_GETITEMDATA, i, 0) == taskId) {
            targetIndex = i;
            break;
        }
    }
    if (targetIndex >= 0) SendMessageW(g_app.settingsTaskCombo, CB_SETCURSEL, targetIndex, 0);

    int sel = (int)SendMessageW(g_app.settingsTaskCombo, CB_GETCURSEL, 0, 0);
    if (sel < 0) {
        SetWindowTextW(g_app.settingsInfoLabel, L"設定できるタスクがありません。タスクタブで追加してください。");
        EnableWindow(g_app.settingsTaskCombo, FALSE);
        EnableWindow(g_app.settingsColorCombo, FALSE);
        EnableWindow(g_app.settingsApplyButton, FALSE);
        SendMessageW(g_app.settingsColorCombo, CB_SETCURSEL, 0, 0);
        return;
    }

    EnableWindow(g_app.settingsTaskCombo, TRUE);
    EnableWindow(g_app.settingsColorCombo, TRUE);
    EnableWindow(g_app.settingsApplyButton, TRUE);

    int id = (int)SendMessageW(g_app.settingsTaskCombo, CB_GETITEMDATA, sel, 0);
    const Task* task = FindTaskByIdConst(id);
    if (!task) {
        SetWindowTextW(g_app.settingsInfoLabel, L"選択したタスクが見つかりません。タスク一覧を更新してください。");
        return;
    }

    int colorSel = task->color == TaskColorMode::White ? 1 : task->color == TaskColorMode::Yellow ? 2 : task->color == TaskColorMode::Orange ? 3 : 0;
    SendMessageW(g_app.settingsColorCombo, CB_SETCURSEL, colorSel, 0);
    std::wstring info = L"対象: " + BuildTaskDisplayName(*task) + L" / 現在: " + TaskColorModeToDisplay(task->color);
    SetWindowTextW(g_app.settingsInfoLabel, info.c_str());
}

void UpdateVisibleTab() {
    bool showTask = (g_app.currentTab == 0);
    ShowControl(g_app.filterCombo, showTask);
    ShowControl(g_app.pendingCheck, showTask);
    ShowControl(g_app.hintLabel, showTask);
    ShowControl(g_app.pathLabel, showTask);
    ShowControl(g_app.list, showTask);
    ShowControl(g_app.statusLabel, showTask);
    if (g_app.emptyLabel) ShowWindow(g_app.emptyLabel, showTask && ListView_GetItemCount(g_app.list) == 0 ? SW_SHOW : SW_HIDE);

    ShowControl(g_app.settingsTaskLabel, !showTask);
    ShowControl(g_app.settingsTaskCombo, !showTask);
    ShowControl(g_app.settingsColorLabel, !showTask);
    ShowControl(g_app.settingsColorCombo, !showTask);
    ShowControl(g_app.settingsApplyButton, !showTask);
    ShowControl(g_app.settingsInfoLabel, !showTask);
    ShowControl(g_app.settingsRulesLabel, !showTask);
}

void SetCurrentTab(int tabIndex) {
    g_app.currentTab = ClampInt(tabIndex, 0, 1);
    if (g_app.tab) TabCtrl_SetCurSel(g_app.tab, g_app.currentTab);
    UpdateVisibleTab();
    SaveConfig();
    if (g_app.currentTab == 1) SyncSettingsPanelFromTask(SelectedTaskId());
}

int SelectedTaskId() {
    if (!g_app.list) return 0;
    int sel = ListView_GetNextItem(g_app.list, -1, LVNI_SELECTED);
    if (sel < 0) return 0;
    LVITEMW item{};
    item.mask = LVIF_PARAM;
    item.iItem = sel;
    if (ListView_GetItem(g_app.list, &item)) {
        return (int)item.lParam;
    }
    return 0;
}

void AutoSizeColumns() {
    if (!g_app.list) return;
    RECT rc{};
    GetClientRect(g_app.list, &rc);
    int width = std::max<int>(240, static_cast<int>(rc.right - rc.left - GetSystemMetrics(SM_CXVSCROLL) - 8));
    const int unitWidth = ScaleForDpi(88, g_app.currentDpi);
    const int dueWidth = ScaleForDpi(220, g_app.currentDpi);
    const int categoryWidth = ScaleForDpi(140, g_app.currentDpi);
    const int notifyWidth = ScaleForDpi(110, g_app.currentDpi);
    const int stateWidth = ScaleForDpi(92, g_app.currentDpi);
    const int enabledWidth = ScaleForDpi(88, g_app.currentDpi);
    const int snoozeWidth = ScaleForDpi(170, g_app.currentDpi);
    int fixed = unitWidth + dueWidth + categoryWidth + notifyWidth + stateWidth + enabledWidth + snoozeWidth;
    int titleWidth = std::max(ScaleForDpi(260, g_app.currentDpi), width - fixed);
    ListView_SetColumnWidth(g_app.list, 0, unitWidth);
    ListView_SetColumnWidth(g_app.list, 1, dueWidth);
    ListView_SetColumnWidth(g_app.list, 2, categoryWidth);
    ListView_SetColumnWidth(g_app.list, 3, titleWidth);
    ListView_SetColumnWidth(g_app.list, 4, notifyWidth);
    ListView_SetColumnWidth(g_app.list, 5, stateWidth);
    ListView_SetColumnWidth(g_app.list, 6, enabledWidth);
    ListView_SetColumnWidth(g_app.list, 7, snoozeWidth);
}

void RefreshList(int preferTaskId) {
    if (!g_app.list) return;
    int keepTaskId = preferTaskId != 0 ? preferTaskId : SelectedTaskId();
    ListView_DeleteAllItems(g_app.list);
    auto now = std::time(nullptr);
    int row = 0;
    for (const auto& task : g_app.tasks) {
        if (g_app.filterIndex == 1 && task.type != TaskType::Yearly) continue;
        if (g_app.filterIndex == 2 && task.type != TaskType::Monthly) continue;
        if (g_app.filterIndex == 3 && task.type != TaskType::Daily) continue;
        if (g_app.onlyPending && IsCompletedForCurrentPeriod(task, now)) continue;

        LVITEMW item{};
        item.mask = LVIF_TEXT | LVIF_PARAM;
        item.iItem = row;
        std::wstring typeText = TaskTypeToDisplay(task.type);
        item.pszText = typeText.data();
        item.lParam = task.id;
        int inserted = ListView_InsertItem(g_app.list, &item);
        if (inserted >= 0) {
            std::wstring dueDisplay = DueDisplay(task);
            ListView_SetItemText(g_app.list, inserted, 1, (LPWSTR)dueDisplay.c_str());
            ListView_SetItemText(g_app.list, inserted, 2, (LPWSTR)task.category.c_str());
            ListView_SetItemText(g_app.list, inserted, 3, (LPWSTR)task.title.c_str());
            ListView_SetItemText(g_app.list, inserted, 4, (LPWSTR)task.cachedNotifyDisplay.c_str());
            std::wstring state = IsCompletedForCurrentPeriod(task, now) ? L"完了" : L"未完了";
            ListView_SetItemText(g_app.list, inserted, 5, (LPWSTR)state.c_str());
            auto enabled = BoolText(task.enabled);
            ListView_SetItemText(g_app.list, inserted, 6, (LPWSTR)enabled.c_str());
            std::wstring snooze = task.snoozedUntil > now ? FormatDateTime(task.snoozedUntil) : L"";
            ListView_SetItemText(g_app.list, inserted, 7, (LPWSTR)snooze.c_str());
            if (keepTaskId != 0 && task.id == keepTaskId) {
                ListView_SetItemState(g_app.list, inserted, LVIS_SELECTED | LVIS_FOCUSED, LVIS_SELECTED | LVIS_FOCUSED);
                ListView_EnsureVisible(g_app.list, inserted, FALSE);
            }
            ++row;
        }
    }
    AutoSizeColumns();
    UpdateStatus();
    PopulateSettingsTaskCombo(keepTaskId);
    SyncSettingsPanelFromTask(keepTaskId);
    UpdateVisibleTab();
}

bool BrowseCsvPath(bool saveMode, std::wstring& selectedPath) {
    wchar_t fileBuf[MAX_PATH * 4]{};
    StringCchCopyW(fileBuf, ARRAYSIZE(fileBuf), selectedPath.c_str());
    OPENFILENAMEW ofn{};
    ofn.lStructSize = sizeof(ofn);
    ofn.hwndOwner = g_app.mainWnd;
    ofn.lpstrFilter = L"CSV ファイル (*.csv)\0*.csv\0すべてのファイル (*.*)\0*.*\0";
    ofn.lpstrFile = fileBuf;
    ofn.nMaxFile = ARRAYSIZE(fileBuf);
    ofn.Flags = OFN_EXPLORER | OFN_PATHMUSTEXIST | OFN_HIDEREADONLY;
    ofn.lpstrDefExt = L"csv";
    BOOL ok = saveMode ? GetSaveFileNameW(&ofn) : GetOpenFileNameW(&ofn);
    if (!ok) return false;
    selectedPath = fileBuf;
    return true;
}

void ShowError(const std::wstring& msg) {
    MessageBoxW(g_app.mainWnd, msg.c_str(), L"デスクトップリマインダー", MB_ICONERROR | MB_OK);
}

void ShowInfo(const std::wstring& msg) {
    MessageBoxW(g_app.mainWnd, msg.c_str(), L"デスクトップリマインダー", MB_ICONINFORMATION | MB_OK);
}

void ShowTrayBalloon(const Task& task) {
    g_app.tray.uFlags = NIF_INFO;
    StringCchCopyW(g_app.tray.szInfoTitle, ARRAYSIZE(g_app.tray.szInfoTitle), L"デスクトップリマインダー");
    std::wstring body = task.category.empty() ? task.title : (task.category + L" / " + task.title);
    StringCchCopyW(g_app.tray.szInfo, ARRAYSIZE(g_app.tray.szInfo), body.c_str());
    g_app.tray.dwInfoFlags = NIIF_INFO;
    if (!Shell_NotifyIconW(NIM_MODIFY, &g_app.tray)) {
        AppendAuditLog(L"NOTIFY_FAILED", L"Tray balloon notification failed for task id=" + std::to_wstring(task.id));
    }
}

INT_PTR CALLBACK PopupDlgProc(HWND dlg, UINT msg, WPARAM wParam, LPARAM lParam) {
    static int s_taskId = 0;
    switch (msg) {
    case WM_INITDIALOG: {
        s_taskId = (int)lParam;
        const Task* task = FindTaskByIdConst(s_taskId);
        if (!task) return FALSE;
        HICON icon = LoadIconW(nullptr, IDI_WARNING);
        if (icon) {
            SendDlgItemMessageW(dlg, IDC_POPUP_ICON, STM_SETICON, (WPARAM)icon, 0);
        }
        SetDlgItemTextW(dlg, IDC_POPUP_TITLE, L"リマインダーの期限です。");
        std::wstring text;
        if (!task->category.empty()) text += task->category + L"\n";
        text += task->title + L"\n" + DueDisplay(*task);
        SetWindowTextW(dlg, L"リマインダー");
        SetDlgItemTextW(dlg, IDC_POPUP_TEXT, text.c_str());
        return TRUE;
    }
    case WM_COMMAND: {
        int id = LOWORD(wParam);
        if (id == IDC_POPUP_COMPLETE || id == IDC_POPUP_SNOOZE5 || id == IDC_POPUP_SNOOZE10 || id == IDC_POPUP_SNOOZE30 || id == IDCANCEL) {
            EndDialog(dlg, id);
            return TRUE;
        }
        break;
    }
    }
    return FALSE;
}

void ApplyPopupResult(int taskId, INT_PTR result) {
    Task* task = FindTaskById(taskId);
    if (!task) return;
    if (result == IDC_POPUP_COMPLETE) {
        PushUndo();
        task->lastCompleted = std::time(nullptr);
        task->snoozedUntil = 0;
        SaveTasks();
        RefreshList(taskId);
    } else if (result == IDC_POPUP_SNOOZE5 || result == IDC_POPUP_SNOOZE10 || result == IDC_POPUP_SNOOZE30) {
        int mins = (result == IDC_POPUP_SNOOZE5) ? 5 : (result == IDC_POPUP_SNOOZE10 ? 10 : 30);
        PushUndo();
        task->snoozedUntil = std::time(nullptr) + mins * 60;
        SaveTasks();
        RefreshList(taskId);
    }
}

void ShowNextPopup() {
    if (g_app.popupOpen) return;
    while (!g_app.popupQueue.empty()) {
        int taskId = g_app.popupQueue.front();
        g_app.popupQueue.pop_front();
        if (!FindTaskByIdConst(taskId)) continue;
        g_app.popupOpen = true;
        INT_PTR result = DialogBoxParamW(g_app.instance, MAKEINTRESOURCEW(IDD_POPUP), g_app.mainWnd, PopupDlgProc, taskId);
        ApplyPopupResult(taskId, result);
        g_app.popupOpen = false;
    }
}

void QueuePopupIfNeeded(int taskId) {
    if (std::find(g_app.popupQueue.begin(), g_app.popupQueue.end(), taskId) == g_app.popupQueue.end()) {
        g_app.popupQueue.push_back(taskId);
        PostMessageW(g_app.mainWnd, WMAPP_SHOW_NEXT_POPUP, 0, 0);
    }
}

bool ShouldNotify(const Task& task, std::time_t now) {
    if (!task.enabled) return false;
    if (task.notify == NotifyType::None) return false;
    if (IsCompletedForCurrentPeriod(task, now)) return false;
    auto due = GetDueTimeForCurrentPeriod(task, now);
    if (!due) return false;
    std::time_t target = *due;
    if (task.snoozedUntil > target) target = task.snoozedUntil;
    if (now < target) return false;
    // notify once per period; re-notify only after a later snooze target
    if (task.lastNotified > 0 && task.lastNotified >= target) {
        return false;
    }
    return true;
}

void CheckReminders() {
    auto now = std::time(nullptr);
    bool changed = false;
    for (auto& task : g_app.tasks) {
        if (!ShouldNotify(task, now)) continue;
        if (task.notify == NotifyType::Tray || task.notify == NotifyType::Both) {
            ShowTrayBalloon(task);
        }
        if (task.notify == NotifyType::Popup || task.notify == NotifyType::Both) {
            QueuePopupIfNeeded(task.id);
        }
        task.lastNotified = now;
        changed = true;
    }
    if (changed) SaveTasks();
}

HWND FindAgentWindow() {
    return FindWindowW(AGENT_WINDOW_CLASS, nullptr);
}

void NotifyAgentReload() {
    HWND agent = FindAgentWindow();
    if (agent) PostMessageW(agent, WMAPP_AGENT_RELOAD, 0, 0);
}

bool LaunchEditorProcess() {
    SHELLEXECUTEINFOW sei{};
    sei.cbSize = sizeof(sei);
    sei.fMask = SEE_MASK_NOCLOSEPROCESS;
    sei.lpFile = g_app.exePath.c_str();
    sei.nShow = SW_SHOWNORMAL;
    sei.lpParameters = L"--editor";
    BOOL ok = ShellExecuteExW(&sei);
    if (ok && sei.hProcess) CloseHandle(sei.hProcess);
    return ok == TRUE;
}

bool EnsureAgentProcess() {
    HWND agent = FindAgentWindow();
    if (agent) return true;
    SHELLEXECUTEINFOW sei{};
    sei.cbSize = sizeof(sei);
    sei.fMask = SEE_MASK_NOCLOSEPROCESS;
    sei.lpFile = g_app.exePath.c_str();
    sei.lpParameters = L"--agent";
    sei.nShow = SW_HIDE;
    BOOL ok = ShellExecuteExW(&sei);
    if (!ok) return false;
    if (sei.hProcess) {
        WaitForInputIdle(sei.hProcess, 3000);
        CloseHandle(sei.hProcess);
    }
    return FindAgentWindow() != nullptr;
}

void HideToTray() {
    if (IsEditorMode()) {
        SaveConfig();
        DestroyWindow(g_app.mainWnd);
        return;
    }
    ShowWindow(g_app.mainWnd, SW_HIDE);
}

void ShowFromTray() {
    if (IsAgentMode()) {
        LaunchEditorProcess();
        return;
    }
    if (!LoadTasks()) {
        AppendAuditLog(L"CSV_RELOAD_FAILED", L"Failed to reload CSV on tray restore: " + g_app.csvPath);
    }
    EnsureMainControls();
    ShowWindow(g_app.mainWnd, SW_SHOWNORMAL);
    SetForegroundWindow(g_app.mainWnd);
    RefreshList();
}

void ExitApplication() {
    SaveConfig();
    g_app.exiting = true;
    if (IsAgentMode()) Shell_NotifyIconW(NIM_DELETE, &g_app.tray);
    DestroyWindow(g_app.mainWnd);
}

std::vector<std::wstring> CollectCategories() {
    std::vector<std::wstring> items;
    for (const auto& task : g_app.tasks) {
        std::wstring category = Trim(task.category);
        if (category.empty()) continue;
        if (std::find(items.begin(), items.end(), category) == items.end()) {
            items.push_back(category);
        }
    }
    std::sort(items.begin(), items.end());
    return items;
}

INT_PTR CALLBACK TaskEditorDlgProc(HWND dlg, UINT msg, WPARAM wParam, LPARAM lParam) {
    auto data = reinterpret_cast<Task*>(GetWindowLongPtrW(dlg, GWLP_USERDATA));
    switch (msg) {
    case WM_INITDIALOG: {
        auto init = reinterpret_cast<TaskEditorInitData*>(lParam);
        auto task = init ? init->task : nullptr;
        if (!task) return FALSE;
        SetWindowLongPtrW(dlg, GWLP_USERDATA, (LONG_PTR)task);
        if (!init->title.empty()) {
            SetWindowTextW(dlg, init->title.c_str());
        }
        HWND type = GetDlgItem(dlg, IDC_TASK_TYPE);
        SendMessageW(type, CB_ADDSTRING, 0, (LPARAM)L"年次");
        SendMessageW(type, CB_ADDSTRING, 0, (LPARAM)L"月次");
        SendMessageW(type, CB_ADDSTRING, 0, (LPARAM)L"日次");
        SendMessageW(type, CB_SETCURSEL, (WPARAM)(task->type == TaskType::Yearly ? 0 : task->type == TaskType::Monthly ? 1 : 2), 0);

        HWND notify = GetDlgItem(dlg, IDC_TASK_NOTIFY);
        SendMessageW(notify, CB_ADDSTRING, 0, (LPARAM)L"ポップアップ");
        SendMessageW(notify, CB_ADDSTRING, 0, (LPARAM)L"通知領域");
        SendMessageW(notify, CB_ADDSTRING, 0, (LPARAM)L"両方");
        SendMessageW(notify, CB_ADDSTRING, 0, (LPARAM)L"なし");
        int notifySel = task->notify == NotifyType::Popup ? 0 : task->notify == NotifyType::Tray ? 1 : task->notify == NotifyType::Both ? 2 : 3;
        SendMessageW(notify, CB_SETCURSEL, notifySel, 0);

        SetDlgItemInt(dlg, IDC_TASK_MONTH, task->month, FALSE);
        SetDlgItemInt(dlg, IDC_TASK_DAY, task->day, FALSE);
        SetDlgItemTextW(dlg, IDC_TASK_TIME, FormatTimeOfDay(task->minutesOfDay).c_str());
        HWND category = GetDlgItem(dlg, IDC_TASK_CATEGORY);
        for (const auto& item : CollectCategories()) {
            SendMessageW(category, CB_ADDSTRING, 0, (LPARAM)item.c_str());
        }
        SetDlgItemTextW(dlg, IDC_TASK_CATEGORY, task->category.c_str());
        SetDlgItemTextW(dlg, IDC_TASK_TITLE, task->title.c_str());
        CheckDlgButton(dlg, IDC_TASK_ENABLED, task->enabled ? BST_CHECKED : BST_UNCHECKED);
        SendMessageW(dlg, WM_COMMAND, MAKEWPARAM(IDC_TASK_TYPE, CBN_SELCHANGE), 0);
        return TRUE;
    }
    case WM_COMMAND: {
        int id = LOWORD(wParam);
        if (id == IDC_TASK_TYPE && HIWORD(wParam) == CBN_SELCHANGE) {
            int sel = (int)SendDlgItemMessageW(dlg, IDC_TASK_TYPE, CB_GETCURSEL, 0, 0);
            bool isYearly = sel == 0;
            bool isMonthly = sel == 1;
            bool isDaily = sel == 2;
            EnableWindow(GetDlgItem(dlg, IDC_TASK_MONTH), isYearly);
            EnableWindow(GetDlgItem(dlg, IDC_TASK_DAY), isYearly || isMonthly);
            EnableWindow(GetDlgItem(dlg, IDC_TASK_TIME), TRUE);
            return TRUE;
        }
        if (id == IDOK) {
            if (!data) return TRUE;
            int sel = (int)SendDlgItemMessageW(dlg, IDC_TASK_TYPE, CB_GETCURSEL, 0, 0);
            data->type = sel == 0 ? TaskType::Yearly : sel == 1 ? TaskType::Monthly : TaskType::Daily;

            wchar_t buf[256]{};
            if (data->type == TaskType::Yearly) {
                BOOL ok1 = FALSE, ok2 = FALSE;
                UINT mo = GetDlgItemInt(dlg, IDC_TASK_MONTH, &ok1, FALSE);
                UINT da = GetDlgItemInt(dlg, IDC_TASK_DAY, &ok2, FALSE);
                if (!ok1 || !ok2 || mo < 1 || mo > 12 || da < 1 || da > 31) {
                    MessageBoxW(dlg, L"年次タスクは月と日を正しく入力してください。", L"入力エラー", MB_OK | MB_ICONWARNING);
                    return TRUE;
                }
                data->month = (int)mo;
                data->day = (int)da;
            } else if (data->type == TaskType::Monthly) {
                BOOL ok = FALSE;
                UINT da = GetDlgItemInt(dlg, IDC_TASK_DAY, &ok, FALSE);
                if (!ok || da < 1 || da > 31) {
                    MessageBoxW(dlg, L"月次タスクは日を正しく入力してください。", L"入力エラー", MB_OK | MB_ICONWARNING);
                    return TRUE;
                }
                data->month = 1;
                data->day = (int)da;
            } else {
                data->month = 1;
                data->day = 1;
            }

            GetDlgItemTextW(dlg, IDC_TASK_TIME, buf, ARRAYSIZE(buf));
            int mod = DEFAULT_REMINDER_HOUR * 60;
            if (!ParseTimeOfDay(buf, mod)) {
                MessageBoxW(dlg, L"時間は HH:MM 形式で入力してください。", L"入力エラー", MB_OK | MB_ICONWARNING);
                return TRUE;
            }
            data->minutesOfDay = mod;

            GetDlgItemTextW(dlg, IDC_TASK_CATEGORY, buf, ARRAYSIZE(buf));
            data->category = Trim(buf);
            GetDlgItemTextW(dlg, IDC_TASK_TITLE, buf, ARRAYSIZE(buf));
            data->title = Trim(buf);
            if (data->title.empty()) {
                MessageBoxW(dlg, L"タスク名を入力してください。", L"入力エラー", MB_OK | MB_ICONWARNING);
                return TRUE;
            }

            int notifySel = (int)SendDlgItemMessageW(dlg, IDC_TASK_NOTIFY, CB_GETCURSEL, 0, 0);
            data->notify = notifySel == 0 ? NotifyType::Popup : notifySel == 1 ? NotifyType::Tray : notifySel == 2 ? NotifyType::Both : NotifyType::None;
            data->enabled = IsDlgButtonChecked(dlg, IDC_TASK_ENABLED) == BST_CHECKED;
            EndDialog(dlg, IDOK);
            return TRUE;
        }
        if (id == IDCANCEL) {
            EndDialog(dlg, IDCANCEL);
            return TRUE;
        }
        break;
    }
    }
    return FALSE;
}

bool ShowTaskEditor(Task& task, bool isNew) {
    TaskEditorInitData init{ &task, isNew ? L"タスク追加" : L"タスク編集" };
    INT_PTR result = DialogBoxParamW(g_app.instance, MAKEINTRESOURCEW(IDD_TASK_EDITOR), g_app.mainWnd, TaskEditorDlgProc, (LPARAM)&init);
    return result == IDOK;
}

void AddTask() {
    Task task;
    task.id = NextTaskId();
    task.type = TaskType::Monthly;
    task.day = 1;
    task.month = 1;
    task.minutesOfDay = DEFAULT_REMINDER_HOUR * 60;
    task.notify = NotifyType::Tray;
    task.color = TaskColorMode::Auto;
    task.enabled = true;
    if (!ShowTaskEditor(task, true)) return;
    NormalizeTask(task);
    RebuildTaskCache(task);
    PushUndo();
    g_app.tasks.push_back(task);
    std::sort(g_app.tasks.begin(), g_app.tasks.end(), [](const Task& a, const Task& b) { return a.id < b.id; });
    SaveTasks();
    RefreshList(task.id);
}

void EditTask(int taskId) {
    Task* task = FindTaskById(taskId);
    if (!task) return;
    Task temp = *task;
    if (!ShowTaskEditor(temp, false)) return;
    NormalizeTask(temp);
    RebuildTaskCache(temp);
    PushUndo();
    *task = temp;
    SaveTasks();
    RefreshList(taskId);
}

void DeleteTask(int taskId) {
    Task* task = FindTaskById(taskId);
    if (!task) return;
    if (MessageBoxW(g_app.mainWnd, L"選択したタスクを削除しますか。", L"確認", MB_YESNO | MB_ICONQUESTION) != IDYES) return;
    PushUndo();
    g_app.tasks.erase(std::remove_if(g_app.tasks.begin(), g_app.tasks.end(), [&](const Task& t) { return t.id == taskId; }), g_app.tasks.end());
    SaveTasks();
    RefreshList();
}

void SetTaskComplete(int taskId, bool complete) {
    Task* task = FindTaskById(taskId);
    if (!task) return;
    PushUndo();
    task->lastCompleted = complete ? std::time(nullptr) : 0;
    if (complete) task->snoozedUntil = 0;
    SaveTasks();
    RefreshList(taskId);
}

void SetTaskEnabled(int taskId, bool enabled) {
    Task* task = FindTaskById(taskId);
    if (!task) return;
    PushUndo();
    task->enabled = enabled;
    SaveTasks();
    RefreshList(taskId);
}

void SnoozeTask(int taskId, int minutes) {
    Task* task = FindTaskById(taskId);
    if (!task) return;
    PushUndo();
    task->snoozedUntil = std::time(nullptr) + minutes * 60;
    SaveTasks();
    RefreshList(taskId);
}

void ClearSnooze(int taskId) {
    Task* task = FindTaskById(taskId);
    if (!task) return;
    PushUndo();
    task->snoozedUntil = 0;
    SaveTasks();
    RefreshList(taskId);
}

void UndoLast() {
    if (g_app.undoStack.empty()) return;
    g_app.tasks = g_app.undoStack.back();
    g_app.undoStack.pop_back();
    RebuildAllTaskCache();
    SaveTasks();
    RefreshList();
}

void ImportCsv() {
    std::wstring path = g_app.csvPath;
    if (!BrowseCsvPath(false, path)) return;
    g_app.csvPath = GetSafeCsvPath(path);
    if (g_app.csvPath != NormalizeCsvPath(path)) {
        ShowError(L"ネットワークパスや特殊デバイスパスの CSV は使用できません。");
        return;
    }
    EnsureCsvExists(g_app.csvPath);
    if (!LoadTasks()) {
        AppendAuditLog(L"CSV_RELOAD_FAILED", L"Failed to import CSV: " + g_app.csvPath);
        ShowError(L"CSV の読込に失敗しました。");
    }
    SaveConfig();
    SetPathLabel();
    RefreshList();
    NotifyAgentReload();
}

void ExportCsv() {
    std::wstring path = g_app.csvPath;
    if (!BrowseCsvPath(true, path)) return;
    std::wstring safePath = GetSafeCsvPath(path);
    if (safePath != NormalizeCsvPath(path)) {
        ShowError(L"ネットワークパスや特殊デバイスパスには保存できません。");
        return;
    }
    std::wstring original = g_app.csvPath;
    g_app.csvPath = safePath;
    if (!SaveTasks()) {
        g_app.csvPath = original;
        ShowError(L"CSV の保存に失敗しました。");
        return;
    }
    SaveConfig();
    SetPathLabel();
    RefreshList();
    NotifyAgentReload();
}

void ShowListContextMenu(POINT screenPt, bool hasSelection) {
    HMENU menu = CreatePopupMenu();
    AppendMenuW(menu, MF_STRING, ID_TASK_ADD, L"新規追加");
    AppendMenuW(menu, MF_SEPARATOR, 0, nullptr);
    if (hasSelection) {
        AppendMenuW(menu, MF_STRING, ID_TASK_EDIT, L"編集");
        AppendMenuW(menu, MF_STRING, ID_TASK_DELETE, L"削除");
        AppendMenuW(menu, MF_SEPARATOR, 0, nullptr);
        AppendMenuW(menu, MF_STRING, ID_TASK_COMPLETE, L"完了にする");
        AppendMenuW(menu, MF_STRING, ID_TASK_INCOMPLETE, L"未完了に戻す");
        AppendMenuW(menu, MF_STRING, ID_TASK_ENABLE, L"有効にする");
        AppendMenuW(menu, MF_STRING, ID_TASK_DISABLE, L"無効にする");
        AppendMenuW(menu, MF_SEPARATOR, 0, nullptr);
        AppendMenuW(menu, MF_STRING, ID_TASK_SNOOZE5, L"5分スヌーズ");
        AppendMenuW(menu, MF_STRING, ID_TASK_SNOOZE10, L"10分スヌーズ");
        AppendMenuW(menu, MF_STRING, ID_TASK_SNOOZE30, L"30分スヌーズ");
        AppendMenuW(menu, MF_STRING, ID_TASK_CLEAR_SNOOZE, L"スヌーズ解除");
        AppendMenuW(menu, MF_SEPARATOR, 0, nullptr);
    }
    AppendMenuW(menu, MF_STRING, ID_TASK_IMPORT, L"CSV読込");
    AppendMenuW(menu, MF_STRING, ID_TASK_EXPORT, L"CSV保存");
    AppendMenuW(menu, MF_SEPARATOR, 0, nullptr);
    AppendMenuW(menu, g_app.undoStack.empty() ? MF_GRAYED : MF_STRING, ID_TASK_UNDO, L"元に戻す");
    SetForegroundWindow(g_app.mainWnd);
    TrackPopupMenu(menu, TPM_RIGHTBUTTON, screenPt.x, screenPt.y, 0, g_app.mainWnd, nullptr);
    DestroyMenu(menu);
}

void HandleCommand(WORD cmd) {
    int selectedId = SelectedTaskId();
    switch (cmd) {
    case ID_TRAY_SHOW: ShowFromTray(); break;
    case ID_TRAY_NEW: if (IsAgentMode()) LaunchEditorProcess(); else AddTask(); break;
    case ID_TRAY_EXIT: ExitApplication(); break;
    case ID_TASK_ADD: AddTask(); break;
    case ID_TASK_EDIT: if (selectedId) EditTask(selectedId); break;
    case ID_TASK_DELETE: if (selectedId) DeleteTask(selectedId); break;
    case ID_TASK_COMPLETE: if (selectedId) SetTaskComplete(selectedId, true); break;
    case ID_TASK_INCOMPLETE: if (selectedId) SetTaskComplete(selectedId, false); break;
    case ID_TASK_ENABLE: if (selectedId) SetTaskEnabled(selectedId, true); break;
    case ID_TASK_DISABLE: if (selectedId) SetTaskEnabled(selectedId, false); break;
    case ID_TASK_SNOOZE5: if (selectedId) SnoozeTask(selectedId, 5); break;
    case ID_TASK_SNOOZE10: if (selectedId) SnoozeTask(selectedId, 10); break;
    case ID_TASK_SNOOZE30: if (selectedId) SnoozeTask(selectedId, 30); break;
    case ID_TASK_CLEAR_SNOOZE: if (selectedId) ClearSnooze(selectedId); break;
    case ID_TASK_IMPORT: ImportCsv(); break;
    case ID_TASK_EXPORT: ExportCsv(); break;
    case ID_TASK_UNDO: UndoLast(); break;
    }
}


void DestroyMainControls() {
    HWND controls[] = { g_app.tab, g_app.filterCombo, g_app.pendingCheck, g_app.hintLabel, g_app.pathLabel, g_app.list, g_app.emptyLabel, g_app.statusLabel,
        g_app.settingsTaskLabel, g_app.settingsTaskCombo, g_app.settingsColorLabel, g_app.settingsColorCombo, g_app.settingsApplyButton, g_app.settingsInfoLabel, g_app.settingsRulesLabel };
    for (HWND &w : controls) {
        if (w) {
            DestroyWindow(w);
            w = nullptr;
        }
    }
    if (g_app.font) { DeleteObject(g_app.font); g_app.font = nullptr; }
    if (g_app.fontBold) { DeleteObject(g_app.fontBold); g_app.fontBold = nullptr; }
    if (g_app.bgBrush) { DeleteObject(g_app.bgBrush); g_app.bgBrush = nullptr; }
    g_app.uiCreated = false;
}

bool EnsureMainControls() {
    if (!g_app.mainWnd) return false;
    if (!g_app.uiCreated) {
        CreateMainControls(g_app.mainWnd);
        LayoutMainControls(g_app.mainWnd);
        RefreshList();
    }
    return g_app.uiCreated;
}

void RecreateUiFonts(UINT dpi) {
    if (dpi < 96) dpi = 96;
    g_app.currentDpi = dpi;
    if (g_app.font) { DeleteObject(g_app.font); g_app.font = nullptr; }
    if (g_app.fontBold) { DeleteObject(g_app.fontBold); g_app.fontBold = nullptr; }

    const int normalHeight = -ScaleForDpi(20, dpi);
    const int boldHeight = -ScaleForDpi(20, dpi);

    g_app.font = CreateFontW(normalHeight, 0, 0, 0, FW_NORMAL, FALSE, FALSE, FALSE, DEFAULT_CHARSET,
        OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, CLEARTYPE_QUALITY, DEFAULT_PITCH | FF_DONTCARE, L"Yu Gothic UI");
    g_app.fontBold = CreateFontW(boldHeight, 0, 0, 0, FW_SEMIBOLD, FALSE, FALSE, FALSE, DEFAULT_CHARSET,
        OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, CLEARTYPE_QUALITY, DEFAULT_PITCH | FF_DONTCARE, L"Yu Gothic UI");

    for (HWND w : { g_app.tab, g_app.filterCombo, g_app.pendingCheck, g_app.hintLabel, g_app.pathLabel, g_app.list, g_app.emptyLabel, g_app.statusLabel,
                    g_app.settingsTaskLabel, g_app.settingsTaskCombo, g_app.settingsColorLabel, g_app.settingsColorCombo,
                    g_app.settingsApplyButton, g_app.settingsInfoLabel, g_app.settingsRulesLabel }) {
        if (w) SendMessageW(w, WM_SETFONT, (WPARAM)g_app.font, TRUE);
    }
    if (g_app.filterCombo) SendMessageW(g_app.filterCombo, WM_SETFONT, (WPARAM)g_app.fontBold, TRUE);
    if (g_app.settingsTaskLabel) SendMessageW(g_app.settingsTaskLabel, WM_SETFONT, (WPARAM)g_app.fontBold, TRUE);
    if (g_app.settingsColorLabel) SendMessageW(g_app.settingsColorLabel, WM_SETFONT, (WPARAM)g_app.fontBold, TRUE);
    if (g_app.list) {
        HWND header = ListView_GetHeader(g_app.list);
        if (header) SendMessageW(header, WM_SETFONT, (WPARAM)g_app.fontBold, TRUE);
    }
    if (g_app.filterCombo) {
        SendMessageW(g_app.filterCombo, CB_SETITEMHEIGHT, (WPARAM)-1, ScaleForDpi(28, dpi));
        SendMessageW(g_app.filterCombo, CB_SETITEMHEIGHT, 0, ScaleForDpi(28, dpi));
    }
    if (g_app.settingsTaskCombo) {
        SendMessageW(g_app.settingsTaskCombo, CB_SETITEMHEIGHT, (WPARAM)-1, ScaleForDpi(28, dpi));
        SendMessageW(g_app.settingsTaskCombo, CB_SETITEMHEIGHT, 0, ScaleForDpi(28, dpi));
    }
    if (g_app.settingsColorCombo) {
        SendMessageW(g_app.settingsColorCombo, CB_SETITEMHEIGHT, (WPARAM)-1, ScaleForDpi(28, dpi));
        SendMessageW(g_app.settingsColorCombo, CB_SETITEMHEIGHT, 0, ScaleForDpi(28, dpi));
    }
}

void CreateMainControls(HWND hwnd) {
    g_app.bgBrush = CreateSolidBrush(RGB(248, 249, 251));
    RecreateUiFonts(GetWindowDpiSafe(hwnd));

    g_app.tab = CreateWindowExW(0, WC_TABCONTROLW, L"",
        WS_CHILD | WS_VISIBLE | WS_CLIPSIBLINGS | WS_TABSTOP,
        ScaleForDpi(12, g_app.currentDpi), ScaleForDpi(8, g_app.currentDpi), ScaleForDpi(930, g_app.currentDpi), ScaleForDpi(590, g_app.currentDpi), hwnd, (HMENU)IDC_MAIN_TAB, g_app.instance, nullptr);
    TCITEMW item{};
    item.mask = TCIF_TEXT;
    item.pszText = const_cast<LPWSTR>(L"タスク");
    TabCtrl_InsertItem(g_app.tab, 0, &item);
    item.pszText = const_cast<LPWSTR>(L"設定");
    TabCtrl_InsertItem(g_app.tab, 1, &item);
    TabCtrl_SetCurSel(g_app.tab, g_app.currentTab);

    g_app.filterCombo = CreateWindowExW(0, WC_COMBOBOXW, nullptr,
        WS_CHILD | WS_VISIBLE | WS_VSCROLL | CBS_DROPDOWNLIST,
        0, 0, 0, 0, hwnd, (HMENU)IDC_FILTER_COMBO, g_app.instance, nullptr);
    SendMessageW(g_app.filterCombo, CB_ADDSTRING, 0, (LPARAM)L"すべて");
    SendMessageW(g_app.filterCombo, CB_ADDSTRING, 0, (LPARAM)L"年次");
    SendMessageW(g_app.filterCombo, CB_ADDSTRING, 0, (LPARAM)L"月次");
    SendMessageW(g_app.filterCombo, CB_ADDSTRING, 0, (LPARAM)L"日次");
    SendMessageW(g_app.filterCombo, CB_SETCURSEL, g_app.filterIndex, 0);

    g_app.pendingCheck = CreateWindowExW(0, L"BUTTON", L"未完了のみ",
        WS_CHILD | WS_VISIBLE | BS_AUTOCHECKBOX,
        0, 0, 0, 0, hwnd, (HMENU)IDC_ONLY_PENDING, g_app.instance, nullptr);
    SendMessageW(g_app.pendingCheck, BM_SETCHECK, g_app.onlyPending ? BST_CHECKED : BST_UNCHECKED, 0);

    g_app.hintLabel = CreateWindowExW(0, L"STATIC", L"右クリックで操作 / ダブルクリックで編集",
        WS_CHILD | WS_VISIBLE, 0, 0, 0, 0, hwnd, nullptr, g_app.instance, nullptr);

    g_app.pathLabel = CreateWindowExW(0, L"STATIC", L"", WS_CHILD | WS_VISIBLE | SS_PATHELLIPSIS,
        0, 0, 0, 0, hwnd, (HMENU)IDC_PATH_LABEL, g_app.instance, nullptr);

    g_app.list = CreateWindowExW(WS_EX_CLIENTEDGE, WC_LISTVIEWW, nullptr,
        WS_CHILD | WS_VISIBLE | LVS_REPORT | LVS_SINGLESEL,
        0, 0, 0, 0, hwnd, (HMENU)IDC_TASK_LIST, g_app.instance, nullptr);
    ListView_SetExtendedListViewStyle(g_app.list, LVS_EX_FULLROWSELECT | LVS_EX_DOUBLEBUFFER | LVS_EX_LABELTIP | LVS_EX_INFOTIP | LVS_EX_GRIDLINES);
    SetWindowTheme(g_app.list, L"Explorer", nullptr);
    ListView_SetBkColor(g_app.list, RGB(255, 255, 255));
    ListView_SetTextBkColor(g_app.list, RGB(255, 255, 255));
    ListView_SetTextColor(g_app.list, RGB(32, 35, 40));

    const wchar_t* headers[] = { L"単位", L"予定", L"カテゴリ", L"タスク", L"通知", L"状態", L"有効", L"スヌーズ" };
    for (int i = 0; i < 8; ++i) {
        LVCOLUMNW col{};
        col.mask = LVCF_TEXT | LVCF_WIDTH | LVCF_SUBITEM;
        col.pszText = const_cast<LPWSTR>(headers[i]);
        col.cx = 100;
        col.iSubItem = i;
        ListView_InsertColumn(g_app.list, i, &col);
    }

    g_app.emptyLabel = CreateWindowExW(0, L"STATIC", L"表示できるタスクがありません。\r\n一覧を右クリックして追加できます。",
        WS_CHILD | SS_CENTER, 0, 0, 0, 0, hwnd, nullptr, g_app.instance, nullptr);

    g_app.statusLabel = CreateWindowExW(0, L"STATIC", L"", WS_CHILD | WS_VISIBLE,
        0, 0, 0, 0, hwnd, (HMENU)IDC_STATUS_LABEL, g_app.instance, nullptr);

    g_app.settingsTaskLabel = CreateWindowExW(0, L"STATIC", L"対象タスク",
        WS_CHILD | WS_VISIBLE, 0, 0, 0, 0, hwnd, (HMENU)IDC_SETTINGS_TASKLBL, g_app.instance, nullptr);
    g_app.settingsTaskCombo = CreateWindowExW(0, WC_COMBOBOXW, nullptr,
        WS_CHILD | WS_VISIBLE | WS_VSCROLL | CBS_DROPDOWNLIST,
        0, 0, 0, 0, hwnd, (HMENU)IDC_SETTINGS_TASK, g_app.instance, nullptr);

    g_app.settingsColorLabel = CreateWindowExW(0, L"STATIC", L"手動色",
        WS_CHILD | WS_VISIBLE, 0, 0, 0, 0, hwnd, (HMENU)IDC_SETTINGS_COLORLBL, g_app.instance, nullptr);
    g_app.settingsColorCombo = CreateWindowExW(0, WC_COMBOBOXW, nullptr,
        WS_CHILD | WS_VISIBLE | WS_VSCROLL | CBS_DROPDOWNLIST,
        0, 0, 0, 0, hwnd, (HMENU)IDC_SETTINGS_COLOR, g_app.instance, nullptr);
    SendMessageW(g_app.settingsColorCombo, CB_ADDSTRING, 0, (LPARAM)L"自動");
    SendMessageW(g_app.settingsColorCombo, CB_ADDSTRING, 0, (LPARAM)L"白");
    SendMessageW(g_app.settingsColorCombo, CB_ADDSTRING, 0, (LPARAM)L"黄色");
    SendMessageW(g_app.settingsColorCombo, CB_ADDSTRING, 0, (LPARAM)L"オレンジ");
    SendMessageW(g_app.settingsColorCombo, CB_SETCURSEL, 0, 0);

    g_app.settingsApplyButton = CreateWindowExW(0, L"BUTTON", L"保存",
        WS_CHILD | WS_VISIBLE | WS_TABSTOP,
        0, 0, 0, 0, hwnd, (HMENU)IDC_SETTINGS_APPLY, g_app.instance, nullptr);

    g_app.settingsInfoLabel = CreateWindowExW(0, L"STATIC", L"",
        WS_CHILD | WS_VISIBLE, 0, 0, 0, 0, hwnd, (HMENU)IDC_SETTINGS_INFO, g_app.instance, nullptr);

    g_app.settingsRulesLabel = CreateWindowExW(0, L"STATIC",
        L"自動色の基準\r\n"
        L"・白: 期限まで余裕あり\r\n"
        L"・黄色: 日次は期限の10分前〜10分後、月次・年次は1日前〜1日後\r\n"
        L"・オレンジ: 日次は10分超過、月次・年次は1日超過\r\n"
        L"※ 手動色を選ぶと自動判定より優先されます。",
        WS_CHILD | WS_VISIBLE, 0, 0, 0, 0, hwnd, (HMENU)IDC_SETTINGS_RULES, g_app.instance, nullptr);

    RecreateUiFonts(g_app.currentDpi);
    SetWindowTextW(g_app.statusLabel, L"読み込み中...");
    SetPathLabel();
    PopulateSettingsTaskCombo();
    SyncSettingsPanelFromTask(SelectedTaskId());
    g_app.uiCreated = true;
    UpdateVisibleTab();
}

void LayoutMainControls(HWND hwnd) {
    if (!g_app.uiCreated) return;
    RECT rc{};
    GetClientRect(hwnd, &rc);
    int width = rc.right - rc.left;
    int height = rc.bottom - rc.top;
    int margin = ScaleForDpi(12, g_app.currentDpi);

    MoveWindow(g_app.tab, margin, margin, std::max(ScaleForDpi(320, g_app.currentDpi), width - margin * 2), std::max(ScaleForDpi(320, g_app.currentDpi), height - margin * 2), TRUE);

    RECT page{};
    GetClientRect(g_app.tab, &page);
    TabCtrl_AdjustRect(g_app.tab, FALSE, &page);
    POINT pagePoints[2] = { { page.left, page.top }, { page.right, page.bottom } };
    MapWindowPoints(g_app.tab, hwnd, pagePoints, 2);
    page.left = pagePoints[0].x;
    page.top = pagePoints[0].y;
    page.right = pagePoints[1].x;
    page.bottom = pagePoints[1].y;
    int pageWidth = page.right - page.left;
    int pageHeight = page.bottom - page.top;

    int innerMargin = ScaleForDpi(12, g_app.currentDpi);
    int topY = page.top + innerMargin;
    int comboW = ScaleForDpi(130, g_app.currentDpi);
    int comboH = ScaleForDpi(34, g_app.currentDpi);
    int checkX = page.left + innerMargin + comboW + ScaleForDpi(14, g_app.currentDpi);
    int checkW = ScaleForDpi(170, g_app.currentDpi);
    int checkH = ScaleForDpi(30, g_app.currentDpi);
    int hintX = checkX + checkW + ScaleForDpi(14, g_app.currentDpi);
    int hintH = ScaleForDpi(24, g_app.currentDpi);
    int pathY = topY + ScaleForDpi(38, g_app.currentDpi);
    int pathH = ScaleForDpi(26, g_app.currentDpi);
    int listTop = pathY + ScaleForDpi(36, g_app.currentDpi);
    int statusHeight = ScaleForDpi(24, g_app.currentDpi);
    int listHeight = std::max(ScaleForDpi(180, g_app.currentDpi), static_cast<int>(pageHeight - (listTop - page.top) - statusHeight - innerMargin - ScaleForDpi(6, g_app.currentDpi)));

    MoveWindow(g_app.filterCombo, page.left + innerMargin, topY, comboW, comboH, TRUE);
    MoveWindow(g_app.pendingCheck, checkX, topY + ScaleForDpi(2, g_app.currentDpi), checkW, checkH, TRUE);
    MoveWindow(g_app.hintLabel, hintX, topY + ScaleForDpi(4, g_app.currentDpi), std::max(ScaleForDpi(200, g_app.currentDpi), static_cast<int>(pageWidth - (hintX - page.left) - innerMargin)), hintH, TRUE);
    MoveWindow(g_app.pathLabel, page.left + innerMargin, pathY, std::max(ScaleForDpi(240, g_app.currentDpi), pageWidth - innerMargin * 2), pathH, TRUE);
    MoveWindow(g_app.list, page.left + innerMargin, listTop, std::max(ScaleForDpi(260, g_app.currentDpi), pageWidth - innerMargin * 2), listHeight, TRUE);
    int emptyWidth = ScaleForDpi(520, g_app.currentDpi);
    int emptyHeight = ScaleForDpi(64, g_app.currentDpi);
    MoveWindow(g_app.emptyLabel, std::max(static_cast<int>(page.left + ScaleForDpi(30, g_app.currentDpi)), static_cast<int>(page.left + (pageWidth - emptyWidth) / 2)), listTop + ScaleForDpi(60, g_app.currentDpi), emptyWidth, emptyHeight, TRUE);
    MoveWindow(g_app.statusLabel, page.left + innerMargin, listTop + listHeight + ScaleForDpi(6, g_app.currentDpi), std::max(ScaleForDpi(220, g_app.currentDpi), pageWidth - innerMargin * 2), statusHeight, TRUE);

    int labelW = ScaleForDpi(90, g_app.currentDpi);
    int valueX = page.left + innerMargin + labelW;
    int contentW = std::max(ScaleForDpi(220, g_app.currentDpi), pageWidth - innerMargin * 2 - labelW);
    int rowH = ScaleForDpi(28, g_app.currentDpi);
    int ctrlH = ScaleForDpi(32, g_app.currentDpi);
    int settingTop = page.top + ScaleForDpi(24, g_app.currentDpi);
    MoveWindow(g_app.settingsTaskLabel, page.left + innerMargin, settingTop + ScaleForDpi(6, g_app.currentDpi), labelW, rowH, TRUE);
    MoveWindow(g_app.settingsTaskCombo, valueX, settingTop, contentW, ctrlH, TRUE);
    MoveWindow(g_app.settingsColorLabel, page.left + innerMargin, settingTop + ScaleForDpi(54, g_app.currentDpi), labelW, rowH, TRUE);
    MoveWindow(g_app.settingsColorCombo, valueX, settingTop + ScaleForDpi(48, g_app.currentDpi), ScaleForDpi(180, g_app.currentDpi), ctrlH, TRUE);
    MoveWindow(g_app.settingsApplyButton, valueX + ScaleForDpi(196, g_app.currentDpi), settingTop + ScaleForDpi(48, g_app.currentDpi), ScaleForDpi(90, g_app.currentDpi), ctrlH, TRUE);
    int infoY = settingTop + ScaleForDpi(98, g_app.currentDpi);
    int infoHeight = ScaleForDpi(30, g_app.currentDpi);
    int rulesY = infoY + infoHeight + ScaleForDpi(10, g_app.currentDpi);
    int availableRulesHeight = static_cast<int>(page.bottom) - innerMargin - rulesY;
    int rulesHeight = std::max(0, availableRulesHeight);
    MoveWindow(g_app.settingsInfoLabel, page.left + innerMargin, infoY, std::max(ScaleForDpi(260, g_app.currentDpi), pageWidth - innerMargin * 2), infoHeight, TRUE);
    MoveWindow(g_app.settingsRulesLabel, page.left + innerMargin, rulesY, std::max(ScaleForDpi(320, g_app.currentDpi), pageWidth - innerMargin * 2), rulesHeight, TRUE);

    AutoSizeColumns();
    UpdateVisibleTab();
}

std::optional<std::time_t> GetNextPeriodDueTime(const Task& task, std::time_t now) {
    reminder_core::ScheduleType scheduleType = reminder_core::ScheduleType::Monthly;
    if (task.type == TaskType::Daily) scheduleType = reminder_core::ScheduleType::Daily;
    else if (task.type == TaskType::Yearly) scheduleType = reminder_core::ScheduleType::Yearly;
    return reminder_core::GetNextPeriodDueTime(scheduleType, task.month, task.day, task.minutesOfDay, now);
}

std::optional<std::time_t> GetNextReminderTarget(const Task& task, std::time_t now) {
    if (!task.enabled || task.notify == NotifyType::None) return std::nullopt;

    if (task.snoozedUntil > now) return task.snoozedUntil;

    if (!IsCompletedForCurrentPeriod(task, now)) {
        auto due = GetDueTimeForCurrentPeriod(task, now);
        if (due) {
            std::time_t target = *due;
            if (task.snoozedUntil > target) target = task.snoozedUntil;
            if (task.lastNotified > 0 && task.lastNotified >= target) {
                return GetNextPeriodDueTime(task, now);
            }
            if (target <= now) return now;
            return target;
        }
    }

    return GetNextPeriodDueTime(task, now);
}

void ScheduleReminderTimer(bool runImmediateCheck) {
    if (!g_app.mainWnd) return;
    KillTimer(g_app.mainWnd, TIMER_REMINDER);
    if (runImmediateCheck) CheckReminders();

    auto now = std::time(nullptr);
    std::optional<std::time_t> nextTarget;
    for (const auto& task : g_app.tasks) {
        auto target = GetNextReminderTarget(task, now);
        if (!target) continue;
        if (!nextTarget || *target < *nextTarget) nextTarget = *target;
    }
    if (!nextTarget) return;

    ULONGLONG waitMs = (*nextTarget <= now) ? 1000ULL : static_cast<ULONGLONG>(*nextTarget - now) * 1000ULL;
    if (waitMs < 1000ULL) waitMs = 1000ULL;
    if (waitMs > MAX_TIMER_WAIT_MS) waitMs = MAX_TIMER_WAIT_MS;
    SetTimer(g_app.mainWnd, TIMER_REMINDER, static_cast<UINT>(waitMs), nullptr);
}

void InitTray(HWND hwnd) {
    g_app.tray = {};
    g_app.tray.cbSize = sizeof(g_app.tray);
    g_app.tray.hWnd = hwnd;
    g_app.tray.uID = 1;
    g_app.tray.uCallbackMessage = WMAPP_TRAY;
    g_app.tray.uFlags = NIF_MESSAGE | NIF_ICON | NIF_TIP;
    g_app.tray.hIcon = LoadIconW(nullptr, IDI_APPLICATION);
    StringCchCopyW(g_app.tray.szTip, ARRAYSIZE(g_app.tray.szTip), L"デスクトップリマインダー");
    Shell_NotifyIconW(NIM_ADD, &g_app.tray);
}

void ShowTrayMenu(HWND hwnd) {
    HMENU menu = CreatePopupMenu();
    AppendMenuW(menu, MF_STRING, ID_TRAY_SHOW, L"開く");
    AppendMenuW(menu, MF_STRING, ID_TRAY_NEW, L"新しい編集ウィンドウ");
    AppendMenuW(menu, MF_SEPARATOR, 0, nullptr);
    AppendMenuW(menu, MF_STRING, ID_TRAY_EXIT, L"終了");
    POINT pt{};
    GetCursorPos(&pt);
    SetForegroundWindow(hwnd);
    TrackPopupMenu(menu, TPM_RIGHTBUTTON, pt.x, pt.y, 0, hwnd, nullptr);
    DestroyMenu(menu);
}

LRESULT CALLBACK MainWndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam) {
    switch (msg) {
    case WM_CREATE:
        if (IsEditorMode()) {
            CreateMainControls(hwnd);
            LayoutMainControls(hwnd);
            RefreshList();
        } else {
            InitTray(hwnd);
            ScheduleReminderTimer(false);
            PostMessageW(hwnd, WMAPP_SHOW_NEXT_POPUP, 0, 0);
        }
        return 0;
    case WM_SIZE:
        if (IsEditorMode()) LayoutMainControls(hwnd);
        return 0;
    case WM_DPICHANGED: {
        g_app.currentDpi = HIWORD(wParam);
        RECT* suggested = reinterpret_cast<RECT*>(lParam);
        if (suggested) {
            SetWindowPos(hwnd, nullptr, suggested->left, suggested->top, suggested->right - suggested->left, suggested->bottom - suggested->top, SWP_NOZORDER | SWP_NOACTIVATE);
        }
        if (IsEditorMode()) {
            RecreateUiFonts(g_app.currentDpi);
            LayoutMainControls(hwnd);
            RefreshList(SelectedTaskId());
        }
        return 0;
    }
    case WM_ERASEBKGND: {
        if (IsEditorMode()) {
            RECT rc{};
            GetClientRect(hwnd, &rc);
            FillRect((HDC)wParam, &rc, g_app.bgBrush ? g_app.bgBrush : (HBRUSH)(COLOR_WINDOW + 1));
            return 1;
        }
        break;
    }
    case WM_CTLCOLORSTATIC: {
        HDC hdc = (HDC)wParam;
        HWND ctl = (HWND)lParam;
        SetBkMode(hdc, TRANSPARENT);
        if (ctl == g_app.hintLabel || ctl == g_app.pathLabel || ctl == g_app.statusLabel) SetTextColor(hdc, RGB(92, 100, 112));
        else SetTextColor(hdc, RGB(40, 44, 52));
        return (INT_PTR)(g_app.bgBrush ? g_app.bgBrush : GetSysColorBrush(COLOR_WINDOW));
    }
    case WM_CTLCOLORBTN: {
        HDC hdc = (HDC)wParam;
        SetBkMode(hdc, TRANSPARENT);
        SetTextColor(hdc, RGB(40, 44, 52));
        return (INT_PTR)(g_app.bgBrush ? g_app.bgBrush : GetSysColorBrush(COLOR_WINDOW));
    }
    case WM_CLOSE:
        HideToTray();
        return 0;
    case WM_DESTROY:
        KillTimer(hwnd, TIMER_REMINDER);
        if (IsAgentMode()) Shell_NotifyIconW(NIM_DELETE, &g_app.tray);
        if (g_app.agentMutex) { CloseHandle(g_app.agentMutex); g_app.agentMutex = nullptr; }
        if (g_app.font) { DeleteObject(g_app.font); g_app.font = nullptr; }
        if (g_app.fontBold) { DeleteObject(g_app.fontBold); g_app.fontBold = nullptr; }
        if (g_app.bgBrush) { DeleteObject(g_app.bgBrush); g_app.bgBrush = nullptr; }
        PostQuitMessage(0);
        return 0;
    case WM_COMMAND: {
        WORD id = LOWORD(wParam);
        WORD code = HIWORD(wParam);
        if (id == IDC_FILTER_COMBO && code == CBN_SELCHANGE) {
            g_app.filterIndex = (int)SendMessageW(g_app.filterCombo, CB_GETCURSEL, 0, 0);
            SaveConfig();
            RefreshList();
            return 0;
        }
        if (id == IDC_ONLY_PENDING && code == BN_CLICKED) {
            g_app.onlyPending = SendMessageW(g_app.pendingCheck, BM_GETCHECK, 0, 0) == BST_CHECKED;
            SaveConfig();
            RefreshList();
            return 0;
        }
        if (id == IDC_SETTINGS_TASK && code == CBN_SELCHANGE) {
            int sel = (int)SendMessageW(g_app.settingsTaskCombo, CB_GETCURSEL, 0, 0);
            int taskId = sel >= 0 ? (int)SendMessageW(g_app.settingsTaskCombo, CB_GETITEMDATA, sel, 0) : 0;
            if (g_app.list && taskId != 0) {
                int count = ListView_GetItemCount(g_app.list);
                for (int i = 0; i < count; ++i) {
                    LVITEMW item{};
                    item.mask = LVIF_PARAM;
                    item.iItem = i;
                    if (ListView_GetItem(g_app.list, &item) && (int)item.lParam == taskId) {
                        ListView_SetItemState(g_app.list, i, LVIS_SELECTED | LVIS_FOCUSED, LVIS_SELECTED | LVIS_FOCUSED);
                        ListView_EnsureVisible(g_app.list, i, FALSE);
                        break;
                    }
                }
            }
            SyncSettingsPanelFromTask(taskId);
            return 0;
        }
        if (id == IDC_SETTINGS_APPLY && code == BN_CLICKED) {
            int taskSel = (int)SendMessageW(g_app.settingsTaskCombo, CB_GETCURSEL, 0, 0);
            if (taskSel < 0) return 0;
            int taskId = (int)SendMessageW(g_app.settingsTaskCombo, CB_GETITEMDATA, taskSel, 0);
            Task* task = FindTaskById(taskId);
            if (!task) return 0;
            int colorSel = (int)SendMessageW(g_app.settingsColorCombo, CB_GETCURSEL, 0, 0);
            TaskColorMode newColor = colorSel == 1 ? TaskColorMode::White : colorSel == 2 ? TaskColorMode::Yellow : colorSel == 3 ? TaskColorMode::Orange : TaskColorMode::Auto;
            if (task->color != newColor) {
                PushUndo();
                task->color = newColor;
                SaveTasks();
                RefreshList(taskId);
            } else {
                SyncSettingsPanelFromTask(taskId);
            }
            return 0;
        }
        HandleCommand(id);
        return 0;
    }
    case WM_NOTIFY: {
        if (!IsEditorMode()) return 0;
        NMHDR* hdr = (NMHDR*)lParam;
        if (hdr->idFrom == IDC_MAIN_TAB && hdr->code == TCN_SELCHANGE) {
            SetCurrentTab(TabCtrl_GetCurSel(g_app.tab));
            LayoutMainControls(hwnd);
            return 0;
        }
        if (hdr->idFrom == IDC_TASK_LIST) {
            if (hdr->code == NM_DBLCLK) {
                int id = SelectedTaskId();
                if (id) EditTask(id);
            } else if (hdr->code == NM_RCLICK) {
                POINT pt{};
                GetCursorPos(&pt);
                ShowListContextMenu(pt, SelectedTaskId() != 0);
            } else if (hdr->code == LVN_ITEMCHANGED) {
                auto* nmlv = reinterpret_cast<NMLISTVIEW*>(lParam);
                if ((nmlv->uChanged & LVIF_STATE) != 0 && (nmlv->uNewState & LVIS_SELECTED) != 0) {
                    SyncSettingsPanelFromTask(SelectedTaskId());
                }
            } else if (hdr->code == NM_CUSTOMDRAW) {
                auto* cd = reinterpret_cast<NMLVCUSTOMDRAW*>(lParam);
                if (cd->nmcd.dwDrawStage == CDDS_PREPAINT) {
                    return CDRF_NOTIFYITEMDRAW;
                }
                if (cd->nmcd.dwDrawStage == CDDS_ITEMPREPAINT) {
                    if (cd->nmcd.dwItemSpec < static_cast<DWORD>(ListView_GetItemCount(g_app.list))) {
                        LVITEMW item{};
                        item.mask = LVIF_PARAM;
                        item.iItem = static_cast<int>(cd->nmcd.dwItemSpec);
                        if (ListView_GetItem(g_app.list, &item)) {
                            const Task* task = FindTaskByIdConst(static_cast<int>(item.lParam));
                            auto now = std::time(nullptr);
                            if ((cd->nmcd.uItemState & CDIS_SELECTED) != 0) {
                                cd->clrTextBk = GetSysColor(COLOR_HIGHLIGHT);
                                cd->clrText = GetSysColor(COLOR_HIGHLIGHTTEXT);
                            } else if (task) {
                                cd->clrTextBk = GetTaskRowBackColor(*task, now, static_cast<int>(cd->nmcd.dwItemSpec));
                                cd->clrText = RGB(32, 35, 40);
                            }
                        }
                    }
                    return CDRF_NEWFONT;
                }
            }
        }
        return 0;
    }
    case WM_TIMER:
        if (IsAgentMode() && wParam == TIMER_REMINDER) {
            CheckReminders();
            ScheduleReminderTimer(false);
        }
        return 0;
    case WMAPP_TRAY:
        if (IsAgentMode()) {
            if (lParam == WM_LBUTTONDBLCLK) {
                ShowFromTray();
            } else if (lParam == WM_RBUTTONUP || lParam == WM_CONTEXTMENU) {
                ShowTrayMenu(hwnd);
            }
        }
        return 0;
    case WMAPP_SHOW_NEXT_POPUP:
        if (IsAgentMode()) ShowNextPopup();
        return 0;
    case WMAPP_AGENT_RELOAD:
        if (IsAgentMode()) {
            LoadConfig();
            if (!LoadTasks()) {
                AppendAuditLog(L"CSV_RELOAD_FAILED", L"Failed to reload CSV on agent signal: " + g_app.csvPath);
            }
            ScheduleReminderTimer(false);
        }
        return 0;
    }
    return DefWindowProcW(hwnd, msg, wParam, lParam);
}

} // namespace

int APIENTRY wWinMain(HINSTANCE hInstance, HINSTANCE, LPWSTR, int nCmdShow) {
    auto setDpiContext = reinterpret_cast<BOOL(WINAPI*)(HANDLE)>(GetProcAddress(GetModuleHandleW(L"user32.dll"), "SetProcessDpiAwarenessContext"));
    if (setDpiContext) {
        setDpiContext(DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2);
    } else {
        SetProcessDPIAware();
    }
    INITCOMMONCONTROLSEX icc{};
    icc.dwSize = sizeof(icc);
    icc.dwICC = ICC_LISTVIEW_CLASSES | ICC_STANDARD_CLASSES;
    InitCommonControlsEx(&icc);

    g_app.instance = hInstance;
    wchar_t exeBuf[MAX_PATH]{};
    GetModuleFileNameW(nullptr, exeBuf, ARRAYSIZE(exeBuf));
    g_app.exePath = exeBuf;

    int argc = 0;
    LPWSTR* argv = CommandLineToArgvW(GetCommandLineW(), &argc);
    for (int i = 1; i < argc; ++i) {
        std::wstring arg = ToLower(argv[i]);
        if (arg == L"--agent") g_app.mode = RunMode::Agent;
        else if (arg == L"--editor") g_app.mode = RunMode::Editor;
    }
    if (argv) LocalFree(argv);

    g_app.iniPath = GetDefaultIniPath();
    std::wstring legacyIniPath = GetLegacyIniPath();
    if (legacyIniPath != g_app.iniPath && GetFileAttributesW(legacyIniPath.c_str()) != INVALID_FILE_ATTRIBUTES &&
        GetFileAttributesW(g_app.iniPath.c_str()) == INVALID_FILE_ATTRIBUTES) {
        EnsureDirectoryForFile(g_app.iniPath);
        CopyFileW(legacyIniPath.c_str(), g_app.iniPath.c_str(), TRUE);
    }
    LoadConfig();
    if (g_app.csvPath.empty()) g_app.csvPath = GetDefaultCsvPath();
    EnsureCsvExists(g_app.csvPath);
    if (!LoadTasks()) {
        AppendAuditLog(L"CSV_RELOAD_FAILED", L"Failed to load CSV at startup: " + g_app.csvPath);
    }

    if (IsEditorMode() && !EnsureAgentProcess()) {
        MessageBoxW(nullptr, L"常駐プロセスを起動できませんでした。", L"デスクトップリマインダー", MB_ICONERROR | MB_OK);
    }

    if (IsAgentMode()) {
        g_app.agentMutex = CreateMutexW(nullptr, TRUE, L"Local\\DesktopReminderAgentSingleton");
        if (!g_app.agentMutex || GetLastError() == ERROR_ALREADY_EXISTS) {
            if (g_app.agentMutex) {
                CloseHandle(g_app.agentMutex);
                g_app.agentMutex = nullptr;
            }
            return 0;
        }
    }

    const wchar_t* CLASS_NAME = IsAgentMode() ? AGENT_WINDOW_CLASS : EDITOR_WINDOW_CLASS;
    WNDCLASSEXW wc{};
    wc.cbSize = sizeof(wc);
    wc.lpfnWndProc = MainWndProc;
    wc.hInstance = hInstance;
    wc.hCursor = LoadCursorW(nullptr, IDC_ARROW);
    wc.hIcon = LoadIconW(nullptr, IDI_APPLICATION);
    wc.hbrBackground = (HBRUSH)(COLOR_WINDOW + 1);
    wc.lpszClassName = CLASS_NAME;
    RegisterClassExW(&wc);

    g_app.currentDpi = GetSystemDpiSafe();
    RECT work{};
    SystemParametersInfoW(SPI_GETWORKAREA, 0, &work, 0);
    int desiredWidth = ScaleForDpi(980, g_app.currentDpi);
    int desiredHeight = ScaleForDpi(620, g_app.currentDpi);
    int width = std::min<int>(desiredWidth, static_cast<int>((work.right - work.left) - ScaleForDpi(40, g_app.currentDpi)));
    int height = std::min<int>(desiredHeight, static_cast<int>((work.bottom - work.top) - ScaleForDpi(40, g_app.currentDpi)));
    int x = static_cast<int>(work.left) + std::max<int>(0, static_cast<int>(((work.right - work.left) - width) / 2));
    int y = static_cast<int>(work.top) + std::max<int>(0, static_cast<int>(((work.bottom - work.top) - height) / 2));

    DWORD style = IsAgentMode() ? WS_OVERLAPPED : WS_OVERLAPPEDWINDOW;
    HWND hwnd = CreateWindowExW(0, CLASS_NAME, IsAgentMode() ? L"DesktopReminderAgent" : L"デスクトップリマインダー",
        style,
        IsAgentMode() ? CW_USEDEFAULT : x, IsAgentMode() ? CW_USEDEFAULT : y, IsAgentMode() ? 0 : width, IsAgentMode() ? 0 : height,
        nullptr, nullptr, hInstance, nullptr);
    if (!hwnd) {
        MessageBoxW(nullptr, L"アプリを起動できませんでした。", L"デスクトップリマインダー", MB_ICONERROR | MB_OK);
        return 1;
    }
    g_app.mainWnd = hwnd;

    if (IsEditorMode()) {
        ShowWindow(hwnd, nCmdShow);
        UpdateWindow(hwnd);
        RefreshList();
    } else {
        ShowWindow(hwnd, SW_HIDE);
        UpdateWindow(hwnd);
        CheckReminders();
        ScheduleReminderTimer(false);
    }

    MSG msg;
    while (GetMessageW(&msg, nullptr, 0, 0)) {
        if (!IsDialogMessageW(hwnd, &msg)) {
            TranslateMessage(&msg);
            DispatchMessageW(&msg);
        }
    }
    return (int)msg.wParam;
}
