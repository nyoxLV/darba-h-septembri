cmake_minimum_required(VERSION 3.20)
project(BridgeTest LANGUAGES CXX)

set(CMAKE_CXX_STANDARD 17)
set(CMAKE_CXX_STANDARD_REQUIRED ON)

# ─── WebView2 SDK (same version as the real app: 1.0.2903.40) ──────────────
set(WEBVIEW2_VERSION "1.0.2903.40")
set(WEBVIEW2_DIR "${CMAKE_BINARY_DIR}/webview2_sdk")

# Reuse the SDK already extracted by the main project (saves a download).
if(NOT EXISTS "${WEBVIEW2_DIR}/build/native/include/WebView2.h" AND
   EXISTS "../cpp/build/webview2_sdk/build/native/include/WebView2.h")
    message(STATUS "Reusing WebView2 SDK from ../cpp/build/webview2_sdk ...")
    execute_process(COMMAND ${CMAKE_COMMAND} -E copy_directory
        "${CMAKE_SOURCE_DIR}/../cpp/build/webview2_sdk" "${WEBVIEW2_DIR}")
endif()

if(NOT EXISTS "${WEBVIEW2_DIR}/build/native/include/WebView2.h")
    message(STATUS "Downloading WebView2 SDK v${WEBVIEW2_VERSION}...")
    file(MAKE_DIRECTORY ${WEBVIEW2_DIR})
    execute_process(COMMAND curl -L --fail -o "${CMAKE_BINARY_DIR}/webview2.zip"
        "https://www.nuget.org/api/v2/package/Microsoft.Web.WebView2/${WEBVIEW2_VERSION}"
        RESULT_VARIABLE dl_status)
    if(NOT dl_status EQUAL 0)
        message(FATAL_ERROR "Failed to download WebView2 SDK (status ${dl_status})")
    endif()
    execute_process(COMMAND powershell.exe -NoProfile -Command "Expand-Archive -LiteralPath '${CMAKE_BINARY_DIR}/webview2.zip' -DestinationPath '${WEBVIEW2_DIR}' -Force"
        RESULT_VARIABLE ex_status)
    if(NOT ex_status EQUAL 0)
        message(FATAL_ERROR "Failed to extract WebView2 SDK")
    endif()
endif()

add_executable(BridgeTest WIN32 src/main.cpp)
target_include_directories(BridgeTest PRIVATE ${WEBVIEW2_DIR}/build/native/include)
# Static loader — no WebView2Loader.dll dependency (same as the real app).
target_link_libraries(BridgeTest PRIVATE
    "${WEBVIEW2_DIR}/build/native/x64/WebView2LoaderStatic.lib"
    ole32 uuid shlwapi)

if(MSVC)
    target_compile_options(BridgeTest PRIVATE /utf-8 /W3)
    # C++17 std::filesystem on MSVC needs the explicit library.
    if(CMAKE_CXX_COMPILER_VERSION VERSION_LESS 19.24)
        target_link_libraries(BridgeTest PRIVATE stdc++)
    endif()
endif()

# Copy the test page next to the exe (webui/test.html).
add_custom_command(TARGET BridgeTest POST_BUILD
    COMMAND ${CMAKE_COMMAND} -E copy_directory
        "${CMAKE_SOURCE_DIR}/webui" "$<TARGET_FILE_DIR:BridgeTest>/webui")
