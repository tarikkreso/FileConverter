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
    
    // Integration menu slots
    void onInstallContextMenu();
    void onRemoveContextMenu();
    void onInstallSendTo();
    void onRemoveSendTo();
    void updateIntegrationMenuState();

private:
    void setupUI();
    void setupMenuBar();
    void addFilesToList(const QStringList &filePaths);
    int findFileRow(const QString &filePath);
    void updateConvertButtonState();
    bool canConvertToFormat(Converter::FileFormat sourceFormat, Converter::FileFormat targetFormat);
    QString formatElapsedTime(qint64 ms);
    QString formatRemainingTime(qint64 ms);

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
    
    // Progress timing
    QElapsedTimer elapsedTimer;
    QTimer *progressTimer;
    
    // Integration menu
    QAction *installContextMenuAction;
    QAction *removeContextMenuAction;
    QAction *installSendToAction;
    QAction *removeSendToAction;
};
#endif // MAINWINDOW_H
