# APEX MACRO v9 — Standalone Bridge Test Harness (no cheat code)

Self-contained WebView2 C++ ↔ JS bridge reproduction. **No HID, no overlay, no weapon
patterns**, only the bridge part that was broken: `chrome.webview.postMessage` from
JS reached C++, but C++ → JS replies (`PostWebMessageAsJson`) seemed to never reach the
page's listener.

## The symptom (as originally reported)

- **Console** (C++ side): every JS message is received, and a reply is queued and flushed
  at top level of the UI message loop via `PostWebMessageAsJson`. No FAILED log.
- **Page** (JS side): the listener is attached at load (`listener=yes`), but after all three
  auto-tests finish, the stats read `recv=0 matched=0 pushes=0`.

## Root cause (FIXED)

The messages were delivered every time; the page threw them away.

`PostWebMessageAsJson` delivers `event.data` **already parsed into a JS object**. That's
documented in `WebView2.idl` ("The `data` property of the event arg is the `webMessage`
string parameter parsed as a JSON string into a JavaScript object"). Only
`PostWebMessageAsString` delivers a string. The listener did:

```js
let m; try { m = JSON.parse(e.data); } catch (_) { log('unparseable message: ...'); return; }
recvCount++;
```

`JSON.parse(object)` coerces the object to `"[object Object]"` and throws. The `catch`
returned *before* `recvCount++`, so `recv`/`matched`/`pushes` stayed at `0` and every
request timed out after 15 s. The page's own console div showed
`unparseable message: [object Object]` for each reply, but `cdp_read.js` only reads the
table and stats, so that line never showed up in `_verify.bat` output.

`--disable-gpu`, `--remote-debugging-port`, threading and re-entrancy are **not**
involved. The original flags work fine.

### Fix (in `webui/test.html`)

```js
window.chrome.webview.addEventListener('message', (e) => {
  recvCount++; ...                       // count delivery before parsing
  let m = e.data;                        // already an object with PostWebMessageAsJson
  if (typeof m === 'string') { try { m = JSON.parse(m); } catch (_) { ...; return; } }
  ...
```

**Porting to the real app:** apply the same change wherever the real web UI (e.g. its
`bridge.js`) does `JSON.parse(e.data)` / `JSON.parse(event.data)` in its
`chrome.webview` `message` listener. Alternatively, switch the C++ side to
`PostWebMessageAsString`. Then `e.data` is a string and the existing `JSON.parse` works.
Don't do both. The tolerant listener above handles either.

### Other harness bugs fixed along the way (`src/main.cpp`)

1. **The HRESULT was never checked.** `FlushResponses()` ignored `PostWebMessageAsJson`'s
   return value and set `hr = S_OK` itself (COM doesn't throw, so the `try/catch` did
   nothing). The earlier "every call returns S_OK" observation was never actually measured.
   The return value is now checked and logged.
2. **Invalid JSON for unknown methods.** The error reply had a doubled quote
   (`"unknown method: X""}`), which `PostWebMessageAsJson` rejects with `E_INVALIDARG`.
   Replies without an `id` were invalid JSON too (`{"id":,...}`).
3. **WebView never sized.** `put_Bounds` was never called, so the WebView wasn't fitted to
   the window. It's now set at creation and on `WM_SIZE`.
4. **Empty console window.** A WIN32-subsystem exe has no CRT stdout after `AllocConsole()`,
   so logs only appeared when `_verify.bat` redirected them to `stdout.log`. stdout is now
   bound to the console when it isn't already redirected.

Use `_verify.bat [args]` to check flag combinations without recompiling:

```bat
_verify.bat                                   :: default = --disable-gpu + port 9333 (real app)
_verify.bat "--remote-debugging-port=9333"    :: CDP only, no GPU disable
```

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

`recv` and `matched` must be non-zero. If they are, the C++→JS path is fixed. Port the
listener change (see **Root cause** above) into the real app's web UI. The C++ side in
`E:\APEX_MACRO_V9\cpp\src\bridge.cpp` only needs the HRESULT check.
