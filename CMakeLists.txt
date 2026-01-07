cmake_minimum_required(VERSION 3.19)
project(FileConverter LANGUAGES CXX)

find_package(Qt6 6.5 REQUIRED COMPONENTS Core Widgets LinguistTools)

qt_standard_project_setup()

add_executable(FileConverter
    src/main.cpp
    src/MainWindow.cpp
    src/MainWindow.h
    src/Converter.h src/Converter.cpp
    src/ContextMenu.h src/ContextMenu.cpp
    src/Dropzone.h src/Dropzone.cpp
)

qt_add_translations(
    TARGETS FileConverter
    TS_FILES FileConverter_en_US.ts
)

target_link_libraries(FileConverter
    PRIVATE
        Qt::Core
        Qt::Widgets
)

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
