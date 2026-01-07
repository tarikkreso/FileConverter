#include "MainWindow.h"

#include <QApplication>
#include <QLocale>
#include <QTranslator>
#include <QSettings>
#include <QCommandLineParser>
#include <QFileInfo>

int main(int argc, char *argv[])
{
    QApplication a(argc, argv);
    
    // Set application metadata for QSettings
    QCoreApplication::setOrganizationName("FileConverter");
    QCoreApplication::setApplicationName("FileConverter");

    // Command line parser for context menu integration
    QCommandLineParser parser;
    parser.setApplicationDescription("Offline File Converter - DOCX/PDF & Image Converter");
    parser.addHelpOption();
    parser.addVersionOption();
    
    QCommandLineOption convertOption(QStringList() << "c" << "convert",
                                     "Convert file to specified format",
                                     "format");
    parser.addOption(convertOption);
    
    parser.addPositionalArgument("file", "File to convert (from context menu)");
    
    parser.process(a);

    QTranslator translator;
    const QStringList uiLanguages = QLocale::system().uiLanguages();
    for (const QString &locale : uiLanguages) {
        const QString baseName = "FileConverter_" + QLocale(locale).name();
        if (translator.load(":/i18n/" + baseName)) {
            a.installTranslator(&translator);
            break;
        }
    }
    
    MainWindow w;
    
    // Handle context menu invocation
    const QStringList args = parser.positionalArguments();
    if (!args.isEmpty()) {
        // File passed from context menu
        QStringList validFiles;
        for (const QString &arg : args) {
            if (QFileInfo::exists(arg)) {
                validFiles << arg;
            }
        }
        if (!validFiles.isEmpty()) {
            w.addFiles(validFiles);
        }
    }
    
    w.show();
    return a.exec();
}
