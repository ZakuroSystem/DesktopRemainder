#pragma once

#include <ctime>
#include <optional>
#include <string>

namespace reminder_core {

enum class ScheduleType {
    Yearly,
    Monthly,
    Daily,
};

struct DueSpec {
    int year = 0;
    int month = 1;
    int day = 1;
    int hour = 9;
    int minute = 0;
};

int DaysInMonth(int year, int month);
int ClampInt(int value, int minimum, int maximum);
bool ParseIntNoThrow(const std::wstring& s, int& value);
std::optional<std::time_t> MakeLocalTime(int year, int month, int day, int hour, int minute, int second = 0);
std::optional<DueSpec> ComputeDueSpecForCurrentPeriod(ScheduleType type, int monthValue, int dayValue, int minutesOfDay, const std::tm& nowTm);
std::optional<std::time_t> GetNextPeriodDueTime(ScheduleType type, int monthValue, int dayValue, int minutesOfDay, std::time_t now);

} // namespace reminder_core
