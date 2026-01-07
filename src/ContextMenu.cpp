#include "ContextMenu.h"
#include <QCoreApplication>
#include <QSettings>
#include <QFileInfo>
#include <QDir>
#include <QDebug>
#include <QStandardPaths>
#include <QFile>

#ifdef Q_OS_WIN
#include <windows.h>
#include <shlobj.h>
#include <objbase.h>
#include <shobjidl.h>
#endif

ContextMenu::ContextMenu() {}

QStringList ContextMenu::supportedExtensions()
{
    return QStringList() << ".docx" << ".pptx" << ".pdf" << ".jpg" << ".jpeg" << ".png" << ".webp" << ".heic" << ".heif";
}

QString ContextMenu::getExecutablePath()
{
    QString exePath = QCoreApplication::applicationFilePath();
    return QDir::toNativeSeparators(exePath);
}

QString ContextMenu::getSendToPath()
{
#ifdef Q_OS_WIN
    wchar_t path[MAX_PATH];
    if (SUCCEEDED(SHGetFolderPathW(nullptr, CSIDL_SENDTO, nullptr, 0, path))) {
        return QString::fromWCharArray(path);
    }
#endif
    return QString();
}

QString ContextMenu::getProgId(const QString &extension)
{
    // Get the ProgID for a file extension from registry
#ifdef Q_OS_WIN
    QString extKey = QString("HKEY_CLASSES_ROOT\\%1").arg(extension);
    QSettings settings(extKey, QSettings::NativeFormat);
    QString progId = settings.value(".").toString();
    return progId.isEmpty() ? QString() : progId;
#else
    Q_UNUSED(extension);
    return QString();
#endif
}

bool ContextMenu::registerShellExtension()
{
#ifdef Q_OS_WIN
    QStringList extensions = supportedExtensions();
    bool allSuccess = true;
    
    for (const QString &ext : extensions) {
        QString formatName;
        if (ext == ".docx") formatName = "Word Document";
        else if (ext == ".pptx") formatName = "PowerPoint";
        else if (ext == ".pdf") formatName = "PDF";
        else if (ext == ".jpg" || ext == ".jpeg") formatName = "JPEG Image";
        else if (ext == ".png") formatName = "PNG Image";
        else if (ext == ".webp") formatName = "WebP Image";
        else if (ext == ".heic" || ext == ".heif") formatName = "HEIC Image";
        
        if (!registerForExtension(ext, formatName)) {
            allSuccess = false;
        }
    }
    
    // Notify Windows of the change
    SHChangeNotify(SHCNE_ASSOCCHANGED, SHCNF_IDLIST, nullptr, nullptr);
    
    return allSuccess;
#else
    return false;
#endif
}

bool ContextMenu::registerForExtension(const QString &extension, const QString &formatName)
{
#ifdef Q_OS_WIN
    QString exePath = getExecutablePath();
    
    // Try to register under the ProgID first (e.g., Word.Document.12 for .docx)
    // This makes it appear in the proper context menu for that file type
    QString progId = getProgId(extension);
    
    // Also register under SystemFileAssociations for a fallback
    QStringList regPaths;
    
    if (!progId.isEmpty()) {
        regPaths << QString("HKEY_CLASSES_ROOT\\%1\\shell\\FileConverter").arg(progId);
    }
    
    // SystemFileAssociations is the recommended approach for per-extension context menus
    regPaths << QString("HKEY_CLASSES_ROOT\\SystemFileAssociations\\%1\\shell\\FileConverter").arg(extension);
    
    bool anySuccess = false;
    
    for (const QString &baseKey : regPaths) {
        QSettings settings(baseKey, QSettings::NativeFormat);
        
        // Set menu text - show appropriate conversion options
        QString menuText = QString("Convert with FileConverter");
        settings.setValue(".", menuText);
        settings.setValue("Icon", exePath);
        
        // Set command
        QString commandKey = baseKey + "\\command";
        QSettings commandSettings(commandKey, QSettings::NativeFormat);
        commandSettings.setValue(".", QString("\"%1\" \"%2\"").arg(exePath, "%1"));
        
        settings.sync();
        commandSettings.sync();
        
        if (settings.status() == QSettings::NoError && 
            commandSettings.status() == QSettings::NoError) {
            anySuccess = true;
        }
    }
    
    return anySuccess;
#else
    Q_UNUSED(extension);
    Q_UNUSED(formatName);
    return false;
#endif
}

bool ContextMenu::unregisterShellExtension()
{
#ifdef Q_OS_WIN
    QStringList extensions = supportedExtensions();
    bool allSuccess = true;
    
    for (const QString &ext : extensions) {
        if (!unregisterForExtension(ext)) {
            allSuccess = false;
        }
    }
    
    // Notify Windows of the change
    SHChangeNotify(SHCNE_ASSOCCHANGED, SHCNF_IDLIST, nullptr, nullptr);
    
    return allSuccess;
#else
    return false;
#endif
}

bool ContextMenu::unregisterForExtension(const QString &extension)
{
#ifdef Q_OS_WIN
    bool anySuccess = false;
    
    // Remove from ProgID
    QString progId = getProgId(extension);
    if (!progId.isEmpty()) {
        QString progIdKey = QString("HKEY_CLASSES_ROOT\\%1\\shell\\FileConverter").arg(progId);
        
        // Delete command subkey first
        QString commandKey = progIdKey + "\\command";
        QSettings commandSettings(commandKey, QSettings::NativeFormat);
        commandSettings.clear();
        commandSettings.sync();
        
        // Then delete the main key
        QSettings settings(progIdKey, QSettings::NativeFormat);
        settings.clear();
        settings.sync();
        anySuccess = true;
    }
    
    // Remove from SystemFileAssociations
    QString sysKey = QString("HKEY_CLASSES_ROOT\\SystemFileAssociations\\%1\\shell\\FileConverter").arg(extension);
    
    QString commandKey = sysKey + "\\command";
    QSettings commandSettings(commandKey, QSettings::NativeFormat);
    commandSettings.clear();
    commandSettings.sync();
    
    QSettings settings(sysKey, QSettings::NativeFormat);
    settings.clear();
    settings.sync();
    
    return true;
#else
    Q_UNUSED(extension);
    return false;
#endif
}

bool ContextMenu::isShellRegistered()
{
#ifdef Q_OS_WIN
    // Check if at least one extension is registered in SystemFileAssociations
    QString testKey = QString("HKEY_CLASSES_ROOT\\SystemFileAssociations\\.docx\\shell\\FileConverter");
    QSettings settings(testKey, QSettings::NativeFormat);
    return settings.contains(".");
#else
    return false;
#endif
}

bool ContextMenu::createSendToShortcut()
{
#ifdef Q_OS_WIN
    QString sendToPath = getSendToPath();
    if (sendToPath.isEmpty()) return false;
    
    QString exePath = getExecutablePath();
    QString shortcutPath = sendToPath + "\\FileConverter.lnk";
    
    // Use COM to create a proper Windows shortcut
    CoInitialize(nullptr);
    
    IShellLinkW *pShellLink = nullptr;
    HRESULT hr = CoCreateInstance(CLSID_ShellLink, nullptr, CLSCTX_INPROC_SERVER, 
                                   IID_IShellLinkW, (void**)&pShellLink);
    
    if (SUCCEEDED(hr)) {
        pShellLink->SetPath(reinterpret_cast<LPCWSTR>(exePath.utf16()));
        pShellLink->SetDescription(L"Convert files with FileConverter");
        
        IPersistFile *pPersistFile = nullptr;
        hr = pShellLink->QueryInterface(IID_IPersistFile, (void**)&pPersistFile);
        
        if (SUCCEEDED(hr)) {
            hr = pPersistFile->Save(reinterpret_cast<LPCOLESTR>(shortcutPath.utf16()), TRUE);
            pPersistFile->Release();
        }
        pShellLink->Release();
    }
    
    CoUninitialize();
    return SUCCEEDED(hr);
#else
    return false;
#endif
}

bool ContextMenu::removeSendToShortcut()
{
#ifdef Q_OS_WIN
    QString sendToPath = getSendToPath();
    if (sendToPath.isEmpty()) return false;
    
    QString shortcutPath = sendToPath + "\\FileConverter.lnk";
    return QFile::remove(shortcutPath);
#else
    return false;
#endif
}

bool ContextMenu::isSendToInstalled()
{
#ifdef Q_OS_WIN
    QString sendToPath = getSendToPath();
    if (sendToPath.isEmpty()) return false;
    
    QString shortcutPath = sendToPath + "\\FileConverter.lnk";
    return QFileInfo::exists(shortcutPath);
#else
    return false;
#endif
}

bool ContextMenu::isRegistered()
{
    return isShellRegistered() || isSendToInstalled();
}
