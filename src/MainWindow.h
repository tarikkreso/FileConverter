#ifndef MAINWINDOW_H
#define MAINWINDOW_H

#include <QMainWindow>
#include <QTableWidget>
#include <QPushButton>
#include <QComboBox>
#include <QProgressBar>
#include <QLabel>
#include <QElapsedTimer>
#include <QTimer>
#include <QMenuBar>
#include <QMenu>
#include <QAction>
#include <QCloseEvent>
#include <QPoint>
#include "Dropzone.h"
#include "Converter.h"

class MainWindow : public QMainWindow
{
    Q_OBJECT

public:
    MainWindow(QWidget *parent = nullptr);
    ~MainWindow();

    // Public method for adding files (used by main.cpp for context menu)
    void addFiles(const QStringList &filePaths);

protected:
    void closeEvent(QCloseEvent *event) override;

private slots:
    void onFilesDropped(const QStringList &filePaths);
    void onAddFilesClicked();
    void onClearClicked();
    void onRemoveSelectedClicked();
    void onConvertClicked();
    void onCancelClicked();
    void onConversionStarted(const QString &filePath);
    void onConversionFinished(const QString &filePath, Converter::ConversionStatus status, const QString &outputPath);
    void onConversionError(const QString &filePath, const QString &errorMessage);
    void onAllConversionsFinished();
    void onFormatChanged(int index);
    void updateProgressTimer();
    void onOpenPreferences();
    void showFileListContextMenu(const QPoint &pos);

#ifdef Q_OS_WIN
    // Integration menu slots (Windows shell integration only)
    void onInstallContextMenu();
    void onRemoveContextMenu();
    void onInstallSendTo();
    void onRemoveSendTo();
    void updateIntegrationMenuState();
#endif

private:
    void setupUI();
    void setupMenuBar();
    void addFilesToList(const QStringList &filePaths);
    int findFileRow(const QString &filePath);
    void updateConvertButtonState();
    bool canConvertToFormat(Converter::FileFormat sourceFormat, Converter::FileFormat targetFormat);
    QString formatElapsedTime(qint64 ms);
    QString formatRemainingTime(qint64 ms);
    void startConversionBatch(const QStringList &filePaths, Converter::FileFormat targetFormat);
    void retryRow(int row);
    void updateProgressAfterItem();
    void loadSettings();
    void saveSettings();

    // UI Components
    Dropzone *dropzone;
    QTableWidget *fileListTable;
    QComboBox *formatSelector;
    QPushButton *addFilesButton;
    QPushButton *convertButton;
    QPushButton *cancelButton;
    QPushButton *clearButton;
    QPushButton *removeButton;
    QPushButton *browseOutputButton;
    QLabel *outputDirLabel;
    QProgressBar *progressBar;
    QLabel *statusLabel;
    QLabel *timeLabel;

    // Conversion
    Converter *converter;
    int totalFiles;
    int processedFiles;
    QString outputDirectory;
    QString lastOutputPath;

    // Per-batch outcome counts, used to build an accurate completion summary
    int successCount = 0;
    int failedCount = 0;
    int unsupportedCount = 0;
    int cancelledCount = 0;

    // Persisted preferences
    bool rememberOutputDirectory = false;
    int jpgQuality = 90;
    int maxParallelConversions = 1;

    // Progress timing
    QElapsedTimer elapsedTimer;
    QTimer *progressTimer;
    
#ifdef Q_OS_WIN
    // Integration menu (Windows shell integration only)
    QAction *installContextMenuAction;
    QAction *removeContextMenuAction;
    QAction *installSendToAction;
    QAction *removeSendToAction;
#endif
};
#endif // MAINWINDOW_H
