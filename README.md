# FileConverter

A compact offline GUI utility to convert documents and images (DOCX/PPTX/PDF, JPG/PNG/WEBP/HEIC) using LibreOffice and ImageMagick.

Prerequisites
- Qt 6 development files
- CMake 3.16+
- C++ toolchain (MSVC on Windows)
- LibreOffice and ImageMagick installed and accessible

Build (Qt Creator)
Open the repository root in Qt Creator (File → Open File or Project and select `CMakeLists.txt`).

- Choose a kit (Qt version and compiler) suitable for your platform.
- Run CMake configure in the IDE, then build and run using the selected kit and build profile (Debug/Release).

Run
- Launch the built executable from Qt Creator or from the build output directory.
- Ensure LibreOffice and ImageMagick are on `PATH` or configured by the application.

Sources of interest
- `src/MainWindow.*` — UI and workflow
- `src/Converter.*` — conversion engine and process control
- `src/ContextMenu.*` — Windows shell helper
