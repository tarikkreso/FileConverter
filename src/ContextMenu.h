#ifndef CONTEXTMENU_H
#define CONTEXTMENU_H

#include <QString>
#include <QStringList>

class ContextMenu
{
public:
    ContextMenu();
    
    // Shell context menu (per-extension, shows only for supported files)
    static bool registerShellExtension();
    static bool unregisterShellExtension();
    static bool isShellRegistered();
    
    // Send To folder shortcut (safer alternative)
    static bool createSendToShortcut();
    static bool removeSendToShortcut();
    static bool isSendToInstalled();
    
    // Combined check
    static bool isRegistered();
    
    static QStringList supportedExtensions();
    
private:
    static bool registerForExtension(const QString &extension, const QString &formatName);
    static bool unregisterForExtension(const QString &extension);
    static QString getExecutablePath();
    static QString getSendToPath();
    static QString getProgId(const QString &extension);
};

#endif // CONTEXTMENU_H
