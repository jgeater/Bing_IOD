// Bing_IOD.cpp : Downloads Bing Image of the Day images to My Pictures folder
//

#include <iostream>
#include <string>
#include <filesystem>
#include <fstream>
#include <chrono>
#include <iomanip>
#include <sstream>
#include <vector>
#include <algorithm>

#define NOMINMAX
#include <Windows.h>
#include <winhttp.h>
#include <ShlObj.h>

#ifdef ERROR
#undef ERROR
#endif

#pragma comment(lib, "winhttp.lib")

namespace fs = std::filesystem;

class Logger {
public:
    enum class Level {
        INFO,
        WARNING,
        ERROR
    };

    static void Log(Level level, const std::string& message) {
        // Rotate if > 500 KB
        constexpr std::uintmax_t kMaxSize = 500 * 1024;
        const std::string logName = "bing_iod.log";
        const std::string bakName = "bing_iod.log.bak";
        try {
            if (fs::exists(logName)) {
                auto size = fs::file_size(logName);
                if (size > kMaxSize) {
                    // Remove old backup if present, then rename current
                    if (fs::exists(bakName)) {
                        fs::remove(bakName);
                    }
                    fs::rename(logName, bakName);
                }
            }
        }
        catch (...) {
            // ignore rotation errors; continue logging
        }

        auto now = std::chrono::system_clock::now();
        auto time = std::chrono::system_clock::to_time_t(now);
        std::tm tm;
        localtime_s(&tm, &time);

        std::ostringstream oss;
        oss << std::put_time(&tm, "%Y-%m-%d %H:%M:%S");
        
        std::string levelStr;
        switch (level) {
            case Level::INFO: levelStr = "INFO"; break;
            case Level::WARNING: levelStr = "WARNING"; break;
            case Level::ERROR: levelStr = "ERROR"; break;
        }

        std::string logMessage = "[" + oss.str() + "] [" + levelStr + "] " + message;
        std::cout << logMessage << std::endl;

        std::ofstream logFile(logName, std::ios::app);
        if (logFile.is_open()) {
            logFile << logMessage << std::endl;
            logFile.close();
        }
    }
};

class BingImageDownloader {
private:
    std::wstring picturesPath;
    std::wstring targetFolder;

    std::wstring GetMyPicturesPath() {
        PWSTR path = nullptr;
        if (SUCCEEDED(SHGetKnownFolderPath(FOLDERID_Pictures, 0, nullptr, &path))) {
            std::wstring result(path);
            CoTaskMemFree(path);
            return result;
        }
        return L"";
    }

    std::string WStringToString(const std::wstring& wstr) {
        if (wstr.empty()) {
            return std::string();
        }
        int size = WideCharToMultiByte(CP_UTF8, 0, wstr.c_str(), -1, nullptr, 0, nullptr, nullptr);
        std::string str(size, 0);
        WideCharToMultiByte(CP_UTF8, 0, wstr.c_str(), -1, &str[0], size, nullptr, nullptr);
        return str.substr(0, size - 1);
    }

    std::wstring StringToWString(const std::string& str) {
        if (str.empty()) {
            return std::wstring();
        }
        int size = MultiByteToWideChar(CP_UTF8, 0, str.c_str(), -1, nullptr, 0);
        std::wstring wstr(size, 0);
        MultiByteToWideChar(CP_UTF8, 0, str.c_str(), -1, &wstr[0], size);
        return wstr.substr(0, size - 1);
    }

    bool DownloadHTTP(const std::wstring& server, const std::wstring& path, std::vector<BYTE>& buffer) {
        HINTERNET hSession = nullptr;
        HINTERNET hConnect = nullptr;
        HINTERNET hRequest = nullptr;
        bool success = false;

        try {
            hSession = WinHttpOpen(L"Bing IOD Downloader/1.0",
                WINHTTP_ACCESS_TYPE_DEFAULT_PROXY,
                WINHTTP_NO_PROXY_NAME,
                WINHTTP_NO_PROXY_BYPASS, 0);

            if (!hSession) {
                Logger::Log(Logger::Level::ERROR, "Failed to open WinHTTP session");
                return false;
            }

            hConnect = WinHttpConnect(hSession, server.c_str(), INTERNET_DEFAULT_HTTPS_PORT, 0);
            if (!hConnect) {
                Logger::Log(Logger::Level::ERROR, "Failed to connect to server");
                WinHttpCloseHandle(hSession);
                return false;
            }

            hRequest = WinHttpOpenRequest(hConnect, L"GET", path.c_str(),
                nullptr, WINHTTP_NO_REFERER,
                WINHTTP_DEFAULT_ACCEPT_TYPES,
                WINHTTP_FLAG_SECURE);

            if (!hRequest) {
                Logger::Log(Logger::Level::ERROR, "Failed to open request");
                WinHttpCloseHandle(hConnect);
                WinHttpCloseHandle(hSession);
                return false;
            }

            if (!WinHttpSendRequest(hRequest, WINHTTP_NO_ADDITIONAL_HEADERS, 0,
                WINHTTP_NO_REQUEST_DATA, 0, 0, 0)) {
                Logger::Log(Logger::Level::ERROR, "Failed to send request");
            } else if (!WinHttpReceiveResponse(hRequest, nullptr)) {
                Logger::Log(Logger::Level::ERROR, "Failed to receive response");
            } else {
                DWORD bytesAvailable = 0;
                buffer.clear();

                do {
                    bytesAvailable = 0;
                    if (!WinHttpQueryDataAvailable(hRequest, &bytesAvailable)) {
                        Logger::Log(Logger::Level::ERROR, "Error querying data");
                        break;
                    }

                    if (bytesAvailable > 0) {
                        std::vector<BYTE> temp(bytesAvailable);
                        DWORD bytesRead = 0;

                        if (WinHttpReadData(hRequest, temp.data(), bytesAvailable, &bytesRead)) {
                            buffer.insert(buffer.end(), temp.begin(), temp.begin() + bytesRead);
                        }
                    }
                } while (bytesAvailable > 0);

                success = !buffer.empty();
            }

            WinHttpCloseHandle(hRequest);
            WinHttpCloseHandle(hConnect);
            WinHttpCloseHandle(hSession);
        } catch (...) {
            Logger::Log(Logger::Level::ERROR, "Exception during HTTP download");
            if (hRequest) {
                WinHttpCloseHandle(hRequest);
            }
            if (hConnect) {
                WinHttpCloseHandle(hConnect);
            }
            if (hSession) {
                WinHttpCloseHandle(hSession);
            }
            return false;
        }

        return success;
    }

    std::string ExtractImageUrl(const std::string& json) {
        size_t urlPos = json.find("\"url\":\"");
        if (urlPos != std::string::npos) {
            urlPos += 7;
            size_t endPos = json.find("\"", urlPos);
            if (endPos != std::string::npos) {
                return json.substr(urlPos, endPos - urlPos);
            }
        }
        return "";
    }

    std::string ExtractImageTitle(const std::string& json) {
        size_t titlePos = json.find("\"title\":\"");
        if (titlePos != std::string::npos) {
            titlePos += 9;
            size_t endPos = json.find("\"", titlePos);
            if (endPos != std::string::npos) {
                return json.substr(titlePos, endPos - titlePos);
            }
        }
        return "";
    }

    std::string SanitizeFilename(const std::string& filename) {
        std::string result = filename;
        const std::string invalid = "<>:\"/\\|?*";
        for (char c : invalid) {
            std::replace(result.begin(), result.end(), c, '_');
        }
        return result;
    }

    std::string ExtractCleanFilenameFromUrl(const std::string& imageUrl, int index) {
        // Take the last path segment
        size_t lastSlash = imageUrl.find_last_of('/');
        std::string segment = (lastSlash != std::string::npos) ? imageUrl.substr(lastSlash + 1) : imageUrl;

        // Strip Bing thumb prefix "th?id="
        const std::string thumbPrefix = "th?id=";
        if (segment.rfind(thumbPrefix, 0) == 0 && segment.size() > thumbPrefix.size()) {
            segment = segment.substr(thumbPrefix.size());
        }

        // Drop query/extra params at '?' or '&'
        size_t qm = segment.find('?');
        size_t amp = segment.find('&');
        size_t cut = std::string::npos;
        if (qm != std::string::npos && amp != std::string::npos) {
            cut = std::min(qm, amp);
        } else if (qm != std::string::npos) {
            cut = qm;
        } else if (amp != std::string::npos) {
            cut = amp;
        }
        if (cut != std::string::npos) {
            segment = segment.substr(0, cut);
        }

        // Strip leading "OHR." if present
        const std::string ohrPrefix = "OHR.";
        if (segment.rfind(ohrPrefix, 0) == 0 && segment.size() > ohrPrefix.size()) {
            segment = segment.substr(ohrPrefix.size());
        }

        segment = SanitizeFilename(segment);

        if (segment.empty()) {
            segment = "bing_image_" + std::to_string(index) + ".jpg";
        }
        return segment;
    }

public:
    BingImageDownloader() {
        picturesPath = GetMyPicturesPath();
        targetFolder = picturesPath;

        try {
            if (!fs::exists(targetFolder)) {
                fs::create_directories(targetFolder);
                Logger::Log(Logger::Level::INFO, "Created folder: " + WStringToString(targetFolder));
            }
        } catch (const std::exception& e) {
            Logger::Log(Logger::Level::ERROR, std::string("Failed to ensure directory exists: ") + e.what());
        }
    }

    bool DownloadImages(int numberOfImages = 8) {
        Logger::Log(Logger::Level::INFO, "Starting Bing Image of the Day download...");
        
        int downloaded = 0;
        int skipped = 0;
        int errors = 0;

        for (int i = 0; i < numberOfImages; i++) {
            try {
                std::wstring apiPath = L"/HPImageArchive.aspx?format=js&idx=" + std::to_wstring(i) + L"&n=1&mkt=en-US";
                std::vector<BYTE> jsonData;

                Logger::Log(Logger::Level::INFO, "Fetching metadata for image " + std::to_string(i + 1) + "...");

                if (!DownloadHTTP(L"www.bing.com", apiPath, jsonData)) {
                    Logger::Log(Logger::Level::ERROR, "Failed to fetch metadata for image " + std::to_string(i + 1));
                    errors++;
                    continue;
                }

                std::string json(jsonData.begin(), jsonData.end());
                std::string imageUrl = ExtractImageUrl(json);

                if (imageUrl.empty()) {
                    Logger::Log(Logger::Level::WARNING, "No image URL found for index " + std::to_string(i));
                    errors++;
                    continue;
                }

                std::string filename = ExtractCleanFilenameFromUrl(imageUrl, i);
                std::wstring fullPath = targetFolder + L"\\" + StringToWString(filename);

                if (fs::exists(fullPath)) {
                    Logger::Log(Logger::Level::INFO, "SKIPPED: " + filename + " (already exists)");
                    skipped++;
                    continue;
                }

                Logger::Log(Logger::Level::INFO, "Downloading: " + filename);

                std::wstring imagePath = StringToWString(imageUrl);
                std::vector<BYTE> imageData;
                if (DownloadHTTP(L"www.bing.com", imagePath, imageData)) {
                    std::ofstream outFile(fullPath, std::ios::binary);
                    if (outFile.is_open()) {
                        outFile.write(reinterpret_cast<const char*>(imageData.data()), imageData.size());
                        outFile.close();
                        Logger::Log(Logger::Level::INFO, "DOWNLOADED: " + filename + " (" + std::to_string(imageData.size() / 1024) + " KB)");
                        downloaded++;
                    } else {
                        Logger::Log(Logger::Level::ERROR, "Failed to save file: " + filename);
                        errors++;
                    }
                } else {
                    Logger::Log(Logger::Level::ERROR, "Failed to download image: " + filename);
                    errors++;
                }
            } catch (const std::exception& e) {
                Logger::Log(Logger::Level::ERROR, std::string("Exception while processing image: ") + e.what());
                errors++;
            }
        }

        Logger::Log(Logger::Level::INFO, "=== Download Summary ===");
        Logger::Log(Logger::Level::INFO, "Downloaded: " + std::to_string(downloaded));
        Logger::Log(Logger::Level::INFO, "Skipped: " + std::to_string(skipped));
        Logger::Log(Logger::Level::INFO, "Errors: " + std::to_string(errors));
        Logger::Log(Logger::Level::INFO, "Target folder: " + WStringToString(targetFolder));

        return errors == 0;
    }
};

int main()
{
    try {
        Logger::Log(Logger::Level::INFO, "Bing Image of the Day Downloader Started");
        
        BingImageDownloader downloader;
        
        bool success = downloader.DownloadImages(8);
        
        Logger::Log(Logger::Level::INFO, success ? "Process completed successfully" : "Process completed with errors");
        
        Sleep(5000);
        
        return success ? 0 : 1;
    } catch (const std::exception& e) {
        Logger::Log(Logger::Level::ERROR, std::string("Fatal error: ") + e.what());
        std::cerr << "Fatal error: " << e.what() << std::endl;
        std::cin.get();
        return 1;
    }
}
