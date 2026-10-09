// ═══════════════════════════════════════════════════════════════════════════
// APEX MACRO v9 — STANDALONE BRIDGE TEST HARNESS (no cheat code)
// ───────────────────────────────────────────────────────────────────────────
// Reproduces ONLY the WebView2 C++ ↔ JS bridge from E:\APEX_MACRO_V9\cpp:
//   * same WebView2 SDK version + static loader lib
//   * same environment/controller creation flow (single env, single controller)
//   * file:/// navigation to a local webui folder
//   * IsWebMessageEnabled(TRUE), JS chrome.webview.postMessage ↔ C++ handler
// Everything is logged with timestamps so the "C++→JS replies never arrive"
// bug can be diagnosed in isolation. Console window shows live stats + log;
// stdout gets the same trace (AllocConsole, like the real app).
// ═══════════════════════════════════════════════════════════════════════════

#define WIN32_LEAN_AND_MEAN
#include <windows.h>
#include <wrl.h>
#include <WebView2.h>
#include <WebView2EnvironmentOptions.h>   // CoreWebView2EnvironmentOptions + MakeAndInitialize (same as real app main.cpp)
#include <string>
#include <cstdio>
#include <cstdarg>
#include <cstring>
#include <cctype>
#include <mutex>
#include <queue>
#include <thread>
#include <atomic>
#include <chrono>
#include <filesystem>

using namespace Microsoft::WRL;   // Callback<T>, MakeAndInitialize<T>, ComPtr (same as real app main.cpp:35)

// ─── Logging (console + stdout, timestamped) ──────────────────────────────
static std::mutex g_logMutex;
static void Log(const char* fmt, ...) {
    SYSTEMTIME st; GetLocalTime(&st);
    char ts[32]; sprintf_s(ts, "%02d:%02d:%02d.%03d ", st.wHour, st.wMinute, st.wSecond, st.wMilliseconds);
    std::lock_guard<std::mutex> lk(g_logMutex);
    printf("%s", ts);
    va_list ap; va_start(ap, fmt); vprintf(fmt, ap); va_end(ap);
    fflush(stdout);
}

// ─── Bridge state ──────────────────────────────────────────────────────────
static HWND g_hwnd = nullptr;
static Microsoft::WRL::ComPtr<ICoreWebView2Controller> g_controller;
static Microsoft::WRL::ComPtr<ICoreWebView2>           g_webview;

// Stats (updated from any thread, read by the UI)
struct BridgeStats {
    std::atomic<int> jsToCpp{0};      // messages received from JS
    std::atomic<int> postedOk{0};     // PostWebMessageAsJson S_OK count
    std::atomic<int> postedFail{0};   // non-S_OK count
} g_stats;

// Deferred reply queue (same design as the real app's fixed bridge.cpp):
// replies are queued from OnWebMessage and flushed at top level of the UI loop.
static std::mutex  g_qMutex;
static std::queue<std::wstring> g_pendingReplies;

#define WM_APP_BRIDGE_FLUSH (WM_APP + 0x1A9)

static void PostReply(const std::string& jsonUtf8);   // fwd

// ─── JSON helpers (minimal, same as real bridge.cpp) ──────────────────────
static std::string JsonStr(const std::string& s) {
    std::string out = "\"";
    for (unsigned char c : s) {
        switch (c) {
            case '"':  out += "\\\""; break;
            case '\\': out += "\\\\"; break;
            default:   if (c >= 0x20 && c < 0x7f) out += (char)c; // non-ascii dropped in test harness
        }
    }
    out += "\"";
    return out;
}

// Minimal JSON string-field extractor for the flat {"id":N,"method":"X","payload":"Y"} protocol.
static bool JsonField(const std::string& json, const char* key, std::string& out) {
    size_t p = json.find("\"" + std::string(key)); if (p == std::string::npos) return false;
    // skip past the closing quote of the key name
    p += 1 + strlen(key);
    while (p < json.size() && json[p] != '"') p++;      // to closing quote
    if (p >= json.size()) return false;
    p++;                                                // past closing quote
    while (p < json.size() && (json[p] == ' ' || json[p] == '\t')) p++;  // spaces before ':'
    if (p < json.size() && json[p] == ':') { p++; }     // the colon
    while (p < json.size() && (json[p] == ' ' || json[p] == '\t' || json[p] == '\r' || json[p] == '\n')) p++;  // spaces after ':'
    if (p >= json.size()) return false;
    if (json[p] == '"') {                       // string value
        size_t e = p + 1; while (e < json.size() && json[e] != '\"') e++;
        out.assign(json, p + 1, e - p - 1);     // unescaped for the test protocol (no escapes used)
        return true;
    } else {                                     // numeric value
        size_t e = p; while (e < json.size() && isdigit((unsigned char)json[e])) e++;
        out.assign(json, p, e - p);
        return !out.empty();
    }
}

// ─── HWID generator — identical algorithm to the real app's GenerateHWID() ──
static std::string GenerateHWID() {
    DWORD serial = 0, maxComp = 0, size = 0, flags = 0;
    // (LPCSTR rootPathName, LPSTR volumeNameBuffer, DWORD volumeNameBufferSize, LPDWORD volumeSerialNumber, ...)
    GetVolumeInformationA("C:\\", nullptr, 0, &serial, &maxComp, &size, nullptr, 0);
    char buf[32]; sprintf_s(buf, "%04X-%04X", (unsigned)(serial >> 16), (unsigned)(serial & 0xFFFF));
    return std::string(buf);
}

// ─── C++ → JS: queue a reply and ask the UI thread to flush it ─────────────
static void PostReply(const std::string& jsonUtf8) {
    Log("[bridge] C++→JS (queued): %s\n", jsonUtf8.c_str());
    int wlen = MultiByteToWideChar(CP_UTF8, 0, jsonUtf8.c_str(), -1, nullptr, 0);
    if (wlen <= 0) return;
    std::wstring w(wlen, L'\0');
    MultiByteToWideChar(CP_UTF8, 0, jsonUtf8.c_str(), -1, &w[0], wlen);
    {
        std::lock_guard<std::mutex> lk(g_qMutex);
        g_pendingReplies.push(std::move(w));
    }
    if (g_hwnd) PostMessageW(g_hwnd, WM_APP_BRIDGE_FLUSH, 0, 0);
}

// Flush queued replies — called from WndProc at top level of the message loop.
static void FlushResponses() {
    std::queue<std::wstring> batch;
    {
        std::lock_guard<std::mutex> lk(g_qMutex);
        std::swap(batch, g_pendingReplies);
    }
    if (!batch.empty()) Log("[bridge] flush: %d message(s) at top level\n", (int)batch.size());
    while (!batch.empty()) {
        const std::wstring& w = batch.front();
        HRESULT hr = E_UNEXPECTED;
        try {
            // NOTE: called OUTSIDE any WebView2 callback — the key difference vs. re-entrant posting.
            g_webview->PostWebMessageAsJson(w.c_str());
            hr = S_OK;
        } catch (...) {}
        if (SUCCEEDED(hr)) {
            g_stats.postedOk++;
        } else {
            g_stats.postedFail++;
            Log("[bridge] PostWebMessageAsJson FAILED: 0x%08X\n", (unsigned)hr);
        }
        batch.pop();
    }
}

// ─── Worker-thread push test: 5 events at 200 ms intervals (fwd) ──────────
static void PushTestWorker();

// ─── JS → C++ handler ──────────────────────────────────────────────────────
static void OnWebMessage(ICoreWebView2* /*sender*/, ICoreWebView2WebMessageReceivedEventArgs* args) {
    LPWSTR msgRaw = nullptr;
    HRESULT hr = args->TryGetWebMessageAsString(&msgRaw);
    if (FAILED(hr) || !msgRaw) return;

    std::string utf8;   // convert wide → narrow UTF-8 (same as real bridge.cpp OnWebMessage)
    {
        int len = WideCharToMultiByte(CP_UTF8, 0, msgRaw, -1, nullptr, 0, nullptr, nullptr);
        if (len > 0) { utf8.resize(len - 1); WideCharToMultiByte(CP_UTF8, 0, msgRaw, -1, &utf8[0], len, nullptr, nullptr); }
    }
    CoTaskMemFree(msgRaw);

    g_stats.jsToCpp++;
        Log("[bridge] JS→C++ #%d: %s\n", g_stats.jsToCpp.load(), utf8.c_str());

        // ── Flat JSON parse for the test protocol (no external deps) ────────
        std::string idStr, method, payload;
        JsonField(utf8, "id", idStr);
        JsonField(utf8, "method", method);
        JsonField(utf8, "payload", payload);

        // ── Test protocol handlers (mirror the real app's GetHWID/Echo) ────
        if (method == "GetHWID") {
            PostReply("{\"id\":" + idStr + ",\"result\":\"" + GenerateHWID() + "\"}");
        } else if (method == "Echo") {
            // payload arrives as a JSON string value; echo it back verbatim.
            PostReply("{\"id\":" + idStr + ",\"result\":" + JsonStr(payload.empty() ? std::string("empty") : payload) + "}");
        } else if (method == "Ping") {
            char ts[64]; SYSTEMTIME st; GetLocalTime(&st);
            sprintf_s(ts, "%02d:%02d:%02d.%03d", st.wHour, st.wMinute, st.wSecond, st.wMilliseconds);
            PostReply("{\"id\":" + idStr + ",\"result\":\"pong " + std::string(ts) + "\"}");
        } else if (method == "PushTest") {
            // Fire a worker thread that pushes 5 events at 200 ms intervals —
            // tests cross-thread PostWebMessageAsJson exactly like the real app's PushEvent.
            Log("[bridge] spawning push-test worker\n");
            std::thread(PushTestWorker).detach();
        } else {
            Log("[bridge] unknown method '%s' — replying with error\n", method.c_str());
            PostReply("{\"id\":" + idStr + ",\"error\":\"unknown method: " + JsonStr(method).substr(1, 60) + "\"}");
        }
}

// ─── Worker-thread push test: 5 events at 200 ms intervals ────────────────
static void PushTestWorker() {
    for (int i = 1; i <= 5; i++) {
        char msg[96]; sprintf_s(msg, "{\"type\":\"push_test\",\"n\":%d}", i);
        Log("[worker] pushing event %d/5\n", i);
        PostReply(msg);   // same path as real app's PushEvent (cross-thread)
        Sleep(200);
    }
}

// ─── Window procedure ──────────────────────────────────────────────────────
static LRESULT CALLBACK WndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam) {
    switch (msg) {
    case WM_APP_BRIDGE_FLUSH:
        FlushResponses();   // top level of the UI loop — never re-entrant
        return 0;
    case WM_CLOSE: DestroyWindow(hwnd); return 0;
    case WM_DESTROY: PostQuitMessage(0); return 0;
    }
    return DefWindowProcW(hwnd, msg, wParam, lParam);
}

// ─── Build file:/// URL to webui/test.html next to the exe ────────────────
static std::wstring GetTestUrl() {
    wchar_t path[MAX_PATH]; GetModuleFileNameW(nullptr, path, MAX_PATH);
    std::wstring p(path);
    size_t slash = p.find_last_of(L"\\/");
    if (slash != std::wstring::npos) p.erase(slash + 1);
    p += L"webui/test.html";
    for (auto& c : p) if (c == L'\\') c = L'/';
    return L"file:///" + p;
}

int WINAPI wWinMain(HINSTANCE hInst, HINSTANCE, LPWSTR, int nCmdShow) {
    AllocConsole();   // console window for logs (same as the real app)
    SetConsoleTitleW(L"APEX MACRO v9 — Bridge Test Harness");
    Log("=================================================\n");
    Log("[harness] APEX MACRO v9 standalone bridge test\n");

    CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED);

    // ── Window (same class-creation pattern as the real app) ───────────────
    WNDCLASSEXW wc = {};
    wc.cbSize = sizeof(wc);
    wc.lpfnWndProc = WndProc;
    wc.hInstance = hInst;
    wc.lpszClassName = L"BridgeTestWindow";   // registered AND created with the same name (the V8/V9 mismatch was a real bug in v9)
    wc.hCursor = LoadCursor(nullptr, IDC_ARROW);
    RegisterClassExW(&wc);

    g_hwnd = CreateWindowExW(0, L"BridgeTestWindow", L"APEX MACRO v9 — Bridge Test (standalone)",
        WS_OVERLAPPEDWINDOW, CW_USEDEFAULT, CW_USEDEFAULT, 1280, 860, nullptr, nullptr, hInst, nullptr);

    // ── WebView2 environment + controller (same flow as the real app) ───────
    wchar_t exePath[MAX_PATH]; GetModuleFileNameW(nullptr, exePath, MAX_PATH);
    std::wstring userDataFolder = std::filesystem::path(exePath).parent_path().wstring() + L"\\webview2_data";
    CreateDirectoryW(userDataFolder.c_str(), nullptr);

    Microsoft::WRL::ComPtr<ICoreWebView2EnvironmentOptions> envOptions;
    if (SUCCEEDED(MakeAndInitialize<CoreWebView2EnvironmentOptions>(&envOptions)) && envOptions) {
        // Default = same args as the real app. Override with BRIDGE_BROWSER_ARGS to isolate variables, e.g.:
        //   set BRIDGE_BROWSER_ARGS=            (no extra flags at all)
        //   set BRIDGE_BROWSER_ARGS=--disable-gpu
        //   set BRIDGE_BROWSER_ARGS=--remote-debugging-port=9333
        const char* argsEnv = getenv("BRIDGE_BROWSER_ARGS");
        std::wstring browserArgs;
        if (argsEnv && *argsEnv) {
            int n = MultiByteToWideChar(CP_UTF8, 0, argsEnv, -1, nullptr, 0);
            browserArgs.resize(n > 0 ? n - 1 : 0);
            if (!browserArgs.empty()) MultiByteToWideChar(CP_UTF8, 0, argsEnv, -1, &browserArgs[0], n);
        } else {
            browserArgs = L"--disable-gpu --remote-debugging-port=9333";   // same as real app
        }
        Log("[harness] AdditionalBrowserArguments: '%ls'\n", browserArgs.c_str());
        envOptions->put_AdditionalBrowserArguments(browserArgs.c_str());
    }

    HRESULT hr = CreateCoreWebView2EnvironmentWithOptions(
        nullptr, userDataFolder.c_str(), envOptions.Get(),
        Callback<ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler>(
            [&](HRESULT result, ICoreWebView2Environment* e) -> HRESULT {
                if (FAILED(result)) { Log("[harness] CreateEnvironment FAILED 0x%08X\n", (unsigned)result); return result; }
                Log("[harness] Environment created\n");

                hr = e->CreateCoreWebView2Controller(g_hwnd,
                    Callback<ICoreWebView2CreateCoreWebView2ControllerCompletedHandler>(
                        [&](HRESULT r, ICoreWebView2Controller* controller) -> HRESULT {
                            if (FAILED(r)) { Log("[harness] CreateController FAILED 0x%08X\n", (unsigned)r); return r; }
                            g_controller = controller;
                            g_controller->get_CoreWebView2(&g_webview);
                            g_controller->put_IsVisible(TRUE);

                            ICoreWebView2Settings* settings = nullptr;
                            if (SUCCEEDED(g_webview->get_Settings(&settings)) && settings) {
                                settings->put_IsScriptEnabled(TRUE);
                                settings->put_AreDefaultScriptDialogsEnabled(TRUE);
                                settings->put_IsWebMessageEnabled(TRUE);   // ← required for chrome.webview.postMessage
                            }

                            g_webview->add_WebMessageReceived(
                                Callback<ICoreWebView2WebMessageReceivedEventHandler>(
                                    [](ICoreWebView2* s, ICoreWebView2WebMessageReceivedEventArgs* a) -> HRESULT {
                                        OnWebMessage(s, a);
                                        return S_OK;
                                    }).Get(), nullptr);

                            std::wstring url = GetTestUrl();
                            Log("[harness] navigating to %ls\n", url.c_str());
                            g_webview->Navigate(url.c_str());
                            return S_OK;
                        }).Get());
                if (FAILED(hr)) Log("[harness] CreateCoreWebView2Controller call FAILED 0x%08X\n", (unsigned)hr);
                return hr;
            }).Get());

    // ── Message loop ────────────────────────────────────────────────────────
    ShowWindow(g_hwnd, nCmdShow);
    UpdateWindow(g_hwnd);
    Log("[harness] message loop started — watch the page + this console\n");

    MSG msg;
    while (GetMessageW(&msg, nullptr, 0, 0)) {
        TranslateMessage(&msg);
        DispatchMessageW(&msg);
    }

    CoUninitialize();
    Log("[harness] exited. stats: jsToCpp=%d postedOk=%d postedFail=%d\n",
        g_stats.jsToCpp.load(), g_stats.postedOk.load(), g_stats.postedFail.load());
    return 0;
}
