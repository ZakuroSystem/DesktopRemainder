#include "reminder_core.h"

#include <algorithm>
#include <cerrno>
#include <climits>
#include <cwchar>

namespace reminder_core {

int DaysInMonth(int year, int month) {
    static const int d[] = {31,28,31,30,31,30,31,31,30,31,30,31};
    if (month < 1 || month > 12) return 30;
    if (month != 2) return d[month - 1];
    bool leap = ((year % 4 == 0) && (year % 100 != 0)) || (year % 400 == 0);
    return leap ? 29 : 28;
}

int ClampInt(int value, int minimum, int maximum) {
    if (value < minimum) return minimum;
    if (value > maximum) return maximum;
    return value;
}

bool ParseIntNoThrow(const std::wstring& s, int& value) {
    size_t start = 0;
    while (start < s.size() && iswspace(s[start])) ++start;
    size_t end = s.size();
    while (end > start && iswspace(s[end - 1])) --end;
    if (start == end) return false;

    std::wstring t = s.substr(start, end - start);
    wchar_t* parsedEnd = nullptr;
    errno = 0;
    long v = wcstol(t.c_str(), &parsedEnd, 10);
    if (errno == ERANGE || parsedEnd == t.c_str() || *parsedEnd != L'\0') return false;
    if (v < INT_MIN || v > INT_MAX) return false;
    value = static_cast<int>(v);
    return true;
}

std::optional<std::time_t> MakeLocalTime(int year, int month, int day, int hour, int minute, int second) {
    if (month < 1 || month > 12) return std::nullopt;
    if (day < 1 || day > DaysInMonth(year, month)) return std::nullopt;
    std::tm tm{};
    tm.tm_year = year - 1900;
    tm.tm_mon = month - 1;
    tm.tm_mday = day;
    tm.tm_hour = hour;
    tm.tm_min = minute;
    tm.tm_sec = second;
    tm.tm_isdst = -1;
    std::time_t t = std::mktime(&tm);
    if (t == static_cast<std::time_t>(-1)) return std::nullopt;
    return t;
}

std::optional<DueSpec> ComputeDueSpecForCurrentPeriod(ScheduleType type, int monthValue, int dayValue, int minutesOfDay, const std::tm& nowTm) {
    DueSpec spec;
    spec.year = nowTm.tm_year + 1900;
    spec.month = nowTm.tm_mon + 1;
    spec.day = nowTm.tm_mday;
    spec.hour = std::clamp(minutesOfDay / 60, 0, 23);
    spec.minute = std::clamp(minutesOfDay % 60, 0, 59);

    if (type == ScheduleType::Daily) return spec;

    if (type == ScheduleType::Monthly) {
        spec.day = std::min(std::max(dayValue, 1), DaysInMonth(spec.year, spec.month));
        return spec;
    }

    spec.month = std::min(std::max(monthValue, 1), 12);
    spec.day = std::min(std::max(dayValue, 1), DaysInMonth(spec.year, spec.month));
    return spec;
}

std::optional<std::time_t> GetNextPeriodDueTime(ScheduleType type, int monthValue, int dayValue, int minutesOfDay, std::time_t now) {
    std::tm tm{};
    localtime_s(&tm, &now);

    int hour = minutesOfDay / 60;
    int minute = minutesOfDay % 60;

    switch (type) {
    case ScheduleType::Daily: {
        std::time_t tomorrowBase = now + 24 * 60 * 60;
        std::tm nextTm{};
        localtime_s(&nextTm, &tomorrowBase);
        return MakeLocalTime(nextTm.tm_year + 1900, nextTm.tm_mon + 1, nextTm.tm_mday, hour, minute, 0);
    }
    case ScheduleType::Monthly: {
        int year = tm.tm_year + 1900;
        int month = tm.tm_mon + 2;
        if (month == 13) {
            month = 1;
            ++year;
        }
        int day = std::min(std::max(dayValue, 1), DaysInMonth(year, month));
        return MakeLocalTime(year, month, day, hour, minute, 0);
    }
    case ScheduleType::Yearly: {
        int year = tm.tm_year + 1901;
        int month = std::min(std::max(monthValue, 1), 12);
        int day = std::min(std::max(dayValue, 1), DaysInMonth(year, month));
        return MakeLocalTime(year, month, day, hour, minute, 0);
    }
    }

    return std::nullopt;
}

} // namespace reminder_core
