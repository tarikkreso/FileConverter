cmake_minimum_required(VERSION 3.19)
project(FileConverter LANGUAGES CXX VERSION 1.0.0)

find_package(Qt6 6.5 REQUIRED COMPONENTS Core Widgets LinguistTools)

qt_standard_project_setup()

set(APP_ICON_MACOS resources/AppIcon.icns)
set(APP_ICON_RESOURCE_WINDOWS resources/app.rc)

add_executable(FileConverter
    src/main.cpp
    src/MainWindow.cpp
    src/MainWindow.h
    src/Converter.h src/Converter.cpp
    src/ContextMenu.h src/ContextMenu.cpp
    src/Dropzone.h src/Dropzone.cpp
    src/ImageCodec.h src/ImageCodec.cpp
    src/PreferencesDialog.h src/PreferencesDialog.cpp
    $<$<PLATFORM_ID:Darwin>:${APP_ICON_MACOS}>
    $<$<PLATFORM_ID:Windows>:${APP_ICON_RESOURCE_WINDOWS}>
)

if(APPLE)
    set_source_files_properties(${APP_ICON_MACOS} PROPERTIES
        MACOSX_PACKAGE_LOCATION "Resources"
    )
    set_target_properties(FileConverter PROPERTIES
        MACOSX_BUNDLE TRUE
        MACOSX_BUNDLE_BUNDLE_NAME "FileConverter"
        MACOSX_BUNDLE_GUI_IDENTIFIER "com.fileconverter.app"
        MACOSX_BUNDLE_BUNDLE_VERSION ${PROJECT_VERSION}
        MACOSX_BUNDLE_SHORT_VERSION_STRING ${PROJECT_VERSION}
        MACOSX_BUNDLE_ICON_FILE AppIcon.icns
        MACOSX_BUNDLE_INFO_PLIST ${CMAKE_SOURCE_DIR}/resources/Info.plist.in
    )
endif()

qt_add_translations(
    TARGETS FileConverter
    TS_FILES FileConverter_en_US.ts
)

target_link_libraries(FileConverter
    PRIVATE
        Qt::Core
        Qt::Widgets
)

# Optional in-process WEBP encode/decode (avoids shelling out to ImageMagick
# for this format). Falls back to ImageMagick at runtime if not found here.
find_library(WEBP_LIBRARY webp)
find_path(WEBP_INCLUDE_DIR webp/decode.h)
if(WEBP_LIBRARY AND WEBP_INCLUDE_DIR)
    target_include_directories(FileConverter PRIVATE ${WEBP_INCLUDE_DIR})
    target_link_libraries(FileConverter PRIVATE ${WEBP_LIBRARY})
    target_compile_definitions(FileConverter PRIVATE HAVE_WEBP=1)
else()
    message(STATUS "libwebp not found - WEBP conversions will require ImageMagick at runtime")
endif()

# Native HEIC decode on macOS (no external dependency needed there).
if(APPLE)
    target_link_libraries(FileConverter
        PRIVATE
            "-framework ImageIO"
            "-framework CoreGraphics"
            "-framework CoreFoundation"
    )
endif()

include(GNUInstallDirs)

install(TARGETS FileConverter
    BUNDLE  DESTINATION .
    RUNTIME DESTINATION ${CMAKE_INSTALL_BINDIR}
    LIBRARY DESTINATION ${CMAKE_INSTALL_LIBDIR}
)

qt_generate_deploy_app_script(
    TARGET FileConverter
    OUTPUT_SCRIPT deploy_script
    NO_UNSUPPORTED_PLATFORM_ERROR
)
install(SCRIPT ${deploy_script})
