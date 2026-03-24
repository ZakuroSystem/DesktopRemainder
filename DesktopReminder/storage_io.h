#pragma once

#include <string>

namespace storage_io {

std::wstring CsvHeader();
std::wstring BuildAuditLogPath(const std::wstring& iniPath, const std::wstring& exeDir);
void AppendAuditLog(const std::wstring& logPath, const std::wstring& eventName, const std::wstring& message);
bool WriteTextFileUtf8Atomic(const std::wstring& path, const std::wstring& text, const std::wstring& logPath);
bool ReadTextFileUtf8(const std::wstring& path, std::wstring& outText);
bool EnsureCsvExists(const std::wstring& path, const std::wstring& logPath);

} // namespace storage_io
