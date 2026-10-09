# APEX MACRO v9 — Standalone Bridge Test Harness (no cheat code)

Self-contained WebView2 C++ ↔ JS bridge reproduction. **No HID, no overlay, no weapon
patterns** — only the exact part that keeps bugging: `chrome.webview.postMessage` from
JS reaches C++, but C++ → JS replies (`PostWebMessageAsJson`) never reach the page's
listener.

## The bug (reproduced in isolation)

Run it and watch both sides:

- **Console** (C++ side): every JS message is received, a reply is queued, flushed at
  top level of the UI message loop via `PostWebMessageAsJson`, which returns `S_OK`
  (no FAILED log). So C++ *thinks* it delivered.
- **Page** (JS side): listener attached at load (`listener=yes`), but after all three
  auto-tests complete, stats read `recv=0 matched=0 pushes=0`. The page never sees a
  single message from C++.

That's the whole bug: JS→C++ works, C++→JS silently drops. No cheat code involved — so
whatever is wrong lives in how WebView2 / this environment delivers posted messages to
the renderer.

## What I've already ruled out (don't re-test these)

1. **Re-entrancy** — replies are NOT posted inside the `WebMessageReceived` callback.
   They're queued and flushed at top level of the message loop (`WM_APP_BRIDGE_FLUSH`).
   Still broken, so it's not a re-entrant-post issue.
2. **Wrong HRESULT** — every `PostWebMessageAsJson` returns `S_OK`. Not an API failure.
3. **Listener timing** — the page attaches its listener at script parse time (before any
   call), and confirms `listener=yes`. The auto-test fires 1 s after load, well after attach.
4. **Multiple controllers / wrong environment** — single env, single controller, one window.

## What's still untested (the likely culprits)

- **`--disable-gpu`** in the browser args. This is a strong suspect: with GPU disabled,
  WebView2 can run in a degraded mode where posted messages to the renderer get dropped or
  delayed. The real app passes `L"--disable-gpu --remote-debugging-port=9333"`.
- **`--remote-debugging-port=9333`** — CDP attach may interfere with message delivery on some builds.

Use `_verify.bat [args]` to A/B test without recompiling:

```bat
_verify.bat                                   :: default = --disable-gpu + port 9333 (real app)
_verify.bat "--remote-debugging-port=9333"    :: CDP only, no GPU disable   <-- try this first
_verify.bat ""                                :: NO extra flags at all       <-- then this
```

If `recv` goes from `0` to a non-zero number when you drop `--disable-gpu`, that's the fix.

## Files

| File | Purpose |
|------|---------|
| `src/main.cpp` | The whole harness: window, WebView2 env+controller (same flow as real app), JS→C++ handler for GetHWID/Echo/Ping/PushTest, deferred reply queue + top-level flush. Console logs every step with timestamps. |
| `webui/test.html` | Auto-test page: attaches listener at load, fires 3 round-trips after 1 s (GetHWID / Echo / Ping), shows a results table + live stats (`recv`, `matched`, `pushes`). Manual buttons for push/spam/ping. |
| `CMakeLists.txt` | Builds `BridgeTest.exe`. Reuses the WebView2 SDK already extracted by the main project at `../cpp/build/webview2_sdk`; falls back to downloading it if absent. Links the static loader lib (same as real app). |
| `build_and_run.bat` | Configure + build Release + launch, stays open (`pause`). First run downloads/extracts the SDK into `build\`. |
| `_rebuild.bat` | Rebuild only (assumes already configured). |
| `_verify.bat [args]` | Launch with optional browser-arg override, wait for CDP, dump page state + console log. For A/B testing flags. |
| `cdp_read.js` | Reads the live page's results table + stats over CDP (port 9333) using Node's built-in WebSocket — no deps. Used by `_verify.bat`. |

## Build & run

```bat
build_and_run.bat        :: one-shot: configure, build, launch, pause
_rebuild.bat             :: rebuild after editing src/main.cpp
_verify.bat [args]       :: A/B test browser flags (see above)
```

Requires the same toolchain as `E:\APEX_MACRO_V9\cpp`: VS BuildTools 18 + CMake. The
WebView2 SDK is pulled from `../cpp/build/webview2_sdk` if present, else downloaded.

## How to confirm a fix works

A working bridge shows, in `_verify.bat` output:

```
"rows":["1 | GetHWID | 1A3E-5745 | OK: 1A3E-5745 | <ms>", ...]
"stats":["recv=3","matched=3","pushes=0 / 5","listener=yes"]
```

`recv` and `matched` must be non-zero. If they are, the C++→JS path is fixed — port that
change back into `E:\APEX_MACRO_V9\cpp\src\bridge.cpp`.
