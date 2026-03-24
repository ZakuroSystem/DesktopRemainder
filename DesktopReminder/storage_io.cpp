#include "storage_io.h"

#include <windows.h>
#include <shlobj.h>
#include <strsafe.h>

#include <ctime>
#include <fstream>

namespace storage_io {
namespace {

bool EnsureDirectoryForFile(const std::wstring& path) {
    size_t pos = path.find_last_of(L"\\/");
    if (pos == std::wstring::npos) return true;
    std::wstring dir = path.substr(0, pos);
    if (dir.empty()) return true;
    SHCreateDirectoryExW(nullptr, dir.c_str(), nullptr);
    DWORD attrs = GetFileAttributesW(dir.c_str());
    return attrs != INVALID_FILE_ATTRIBUTES && (attrs & FILE_ATTRIBUTE_DIRECTORY);
}

std::string WideToUtf8(const std::wstring& s) {
    if (s.empty()) return "";
    int len = WideCharToMultiByte(CP_UTF8, 0, s.data(), (int)s.size(), nullptr, 0, nullptr, nullptr);
    std::string u(len, '\0');
    WideCharToMultiByte(CP_UTF8, 0, s.data(), (int)s.size(), u.data(), len, nullptr, nullptr);
    return u;
}

std::wstring Utf8ToWide(const std::string& s) {
    if (s.empty()) return L"";
    int len = MultiByteToWideChar(CP_UTF8, 0, s.data(), (int)s.size(), nullptr, 0);
    std::wstring w(len, L'\0');
    MultiByteToWideChar(CP_UTF8, 0, s.data(), (int)s.size(), w.data(), len);
    return w;
}

std::wstring GetBackupPath(const std::wstring& path) {
    return path + L".bak";
}

std::wstring GetPreviousBackupPath(const std::wstring& path) {
    return path + L".bak.prev";
}

} // namespace

std::wstring CsvHeader() {
    return L"id,type,month,day,time,category,title,notify,color,enabled,last_completed,snoozed_until,last_notified\n";
}

std::wstring BuildAuditLogPath(const std::wstring& iniPath, const std::wstring& exeDir) {
    size_t pos = iniPath.find_last_of(L"\\/");
    if (pos != std::wstring::npos) {
        return iniPath.substr(0, pos) + L"\\DesktopReminder.log";
    }
    return exeDir + L"\\DesktopReminder.log";
}

void AppendAuditLog(const std::wstring& logPath, const std::wstring& eventName, const std::wstring& message) {
    if (!EnsureDirectoryForFile(logPath)) return;
    auto now = std::time(nullptr);
    std::tm tm{};
    localtime_s(&tm, &now);
    wchar_t ts[32]{};
    StringCchPrintfW(ts, ARRAYSIZE(ts), L"%04d-%02d-%02d %02d:%02d:%02d",
        tm.tm_year + 1900, tm.tm_mon + 1, tm.tm_mday, tm.tm_hour, tm.tm_min, tm.tm_sec);
    std::wstring line = std::wstring(ts) + L" [" + eventName + L"] " + message + L"\n";
    std::ofstream out(logPath, std::ios::binary | std::ios::app);
    if (!out.is_open()) return;
    auto bytes = WideToUtf8(line);
    out.write(bytes.data(), (std::streamsize)bytes.size());
}

bool WriteTextFileUtf8Atomic(const std::wstring& path, const std::wstring& text, const std::wstring& logPath) {
    if (!EnsureDirectoryForFile(path)) return false;
    std::wstring tempPath = path + L".tmp";
    wchar_t uniqueSuffix[48]{};
    StringCchPrintfW(uniqueSuffix, ARRAYSIZE(uniqueSuffix), L".%lu.%llu", GetCurrentProcessId(), GetTickCount64());
    tempPath += uniqueSuffix;

    std::ofstream out(tempPath, std::ios::binary | std::ios::trunc);
    if (!out.is_open()) return false;
    const unsigned char bom[] = {0xEF, 0xBB, 0xBF};
    out.write((const char*)bom, 3);
    auto bytes = WideToUtf8(text);
    out.write(bytes.data(), (std::streamsize)bytes.size());
    if (!out.good()) {
        out.close();
        DeleteFileW(tempPath.c_str());
        return false;
    }
    out.close();

    DWORD attrs = GetFileAttributesW(path.c_str());
    if (attrs != INVALID_FILE_ATTRIBUTES) {
        std::wstring backupPath = GetBackupPath(path);
        std::wstring previousBackupPath = GetPreviousBackupPath(path);
        DWORD backupAttrs = GetFileAttributesW(backupPath.c_str());
        if (backupAttrs != INVALID_FILE_ATTRIBUTES) {
            if (!MoveFileExW(backupPath.c_str(), previousBackupPath.c_str(), MOVEFILE_REPLACE_EXISTING | MOVEFILE_WRITE_THROUGH)) {
                AppendAuditLog(logPath, L"BACKUP_ROTATE_FAILED", L"Backup rotate failed: " + backupPath + L" -> " + previousBackupPath);
            }
        }
        if (!CopyFileW(path.c_str(), backupPath.c_str(), FALSE)) {
            AppendAuditLog(logPath, L"BACKUP_COPY_FAILED", L"CSV backup copy failed: " + path + L" -> " + backupPath);
        }
    }

    if (!MoveFileExW(tempPath.c_str(), path.c_str(), MOVEFILE_REPLACE_EXISTING | MOVEFILE_WRITE_THROUGH)) {
        DeleteFileW(tempPath.c_str());
        return false;
    }
    return true;
}

bool ReadTextFileUtf8(const std::wstring& path, std::wstring& outText) {
    std::ifstream in(path, std::ios::binary);
    if (!in.is_open()) return false;
    std::string bytes((std::istreambuf_iterator<char>(in)), std::istreambuf_iterator<char>());
    if (bytes.size() >= 3 && (unsigned char)bytes[0] == 0xEF && (unsigned char)bytes[1] == 0xBB && (unsigned char)bytes[2] == 0xBF) {
        bytes.erase(0, 3);
    }
    outText = Utf8ToWide(bytes);
    return true;
}

bool EnsureCsvExists(const std::wstring& path, const std::wstring& logPath) {
    DWORD attrs = GetFileAttributesW(path.c_str());
    if (attrs != INVALID_FILE_ATTRIBUTES) return true;
    bool ok = WriteTextFileUtf8Atomic(path, CsvHeader(), logPath);
    if (!ok) {
        AppendAuditLog(logPath, L"CSV_INIT_FAILED", L"Failed to initialize CSV: " + path);
    }
    return ok;
}

} // namespace storage_io
