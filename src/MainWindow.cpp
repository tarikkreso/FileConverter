#include "MainWindow.h"
#include "ContextMenu.h"
#include <QVBoxLayout>
#include <QHBoxLayout>
#include <QHeaderView>
#include <QFileDialog>
#include <QMessageBox>
#include <QFileInfo>
#include <QGroupBox>
#include <QStatusBar>
#include <QDesktopServices>
#include <QUrl>
#include <QDir>

MainWindow::MainWindow(QWidget *parent)
    : QMainWindow(parent), totalFiles(0), processedFiles(0)
{
    setupUI();
    
    // Create converter
    converter = new Converter(this);
    connect(converter, &Converter::conversionStarted, this, &MainWindow::onConversionStarted);
    connect(converter, &Converter::conversionFinished, this, &MainWindow::onConversionFinished);
    connect(converter, &Converter::conversionError, this, &MainWindow::onConversionError);
    connect(converter, &Converter::allConversionsFinished, this, &MainWindow::onAllConversionsFinished);
    
    // Progress timer for time estimates
    progressTimer = new QTimer(this);
    connect(progressTimer, &QTimer::timeout, this, &MainWindow::updateProgressTimer);
    
    // Set default output directory to user's Documents
    outputDirectory = QDir::homePath() + "/Documents/FileConverter_Output";
}

MainWindow::~MainWindow()
{
}

void MainWindow::addFiles(const QStringList &filePaths)
{
    addFilesToList(filePaths);
}

void MainWindow::setupUI()
{
    setWindowTitle("FileConverter - Offline Document & Image Converter");
    resize(900, 650);
    
    // Setup menu bar
    setupMenuBar();

    // Central widget
    QWidget *centralWidget = new QWidget(this);
    setCentralWidget(centralWidget);

    QVBoxLayout *mainLayout = new QVBoxLayout(centralWidget);
    mainLayout->setContentsMargins(15, 15, 15, 15);
    mainLayout->setSpacing(10);

    // Drop zone
    dropzone = new Dropzone(this);
    connect(dropzone, &Dropzone::filesDropped, this, &MainWindow::onFilesDropped);
    mainLayout->addWidget(dropzone);

    // File list group
    QGroupBox *fileListGroup = new QGroupBox("Files to Convert", this);
    QVBoxLayout *fileListLayout = new QVBoxLayout(fileListGroup);

    // File table
    fileListTable = new QTableWidget(this);
    fileListTable->setColumnCount(4);
    fileListTable->setHorizontalHeaderLabels({"File Name", "Path", "Format", "Status"});
    fileListTable->horizontalHeader()->setStretchLastSection(true);
    fileListTable->horizontalHeader()->setSectionResizeMode(0, QHeaderView::ResizeToContents);
    fileListTable->horizontalHeader()->setSectionResizeMode(1, QHeaderView::Stretch);
    fileListTable->horizontalHeader()->setSectionResizeMode(2, QHeaderView::ResizeToContents);
    fileListTable->setSelectionBehavior(QAbstractItemView::SelectRows);
    fileListTable->setEditTriggers(QAbstractItemView::NoEditTriggers);
    fileListLayout->addWidget(fileListTable);

    mainLayout->addWidget(fileListGroup);

    // Control panel
    QHBoxLayout *controlLayout = new QHBoxLayout();
    
    addFilesButton = new QPushButton("Add Files...", this);
    connect(addFilesButton, &QPushButton::clicked, this, &MainWindow::onAddFilesClicked);
    controlLayout->addWidget(addFilesButton);

    removeButton = new QPushButton("Remove Selected", this);
    connect(removeButton, &QPushButton::clicked, this, &MainWindow::onRemoveSelectedClicked);
    controlLayout->addWidget(removeButton);

    clearButton = new QPushButton("Clear All", this);
    connect(clearButton, &QPushButton::clicked, this, &MainWindow::onClearClicked);
    controlLayout->addWidget(clearButton);

    controlLayout->addStretch();

    QLabel *formatLabel = new QLabel("Convert to:", this);
    controlLayout->addWidget(formatLabel);

    formatSelector = new QComboBox(this);
    formatSelector->addItem("PDF", static_cast<int>(Converter::FileFormat::PDF));
    formatSelector->addItem("DOCX", static_cast<int>(Converter::FileFormat::DOCX));
    formatSelector->addItem("PPTX", static_cast<int>(Converter::FileFormat::PPTX));
    formatSelector->addItem("JPG", static_cast<int>(Converter::FileFormat::JPG));
    formatSelector->addItem("PNG", static_cast<int>(Converter::FileFormat::PNG));
    formatSelector->addItem("WEBP", static_cast<int>(Converter::FileFormat::WEBP));
    formatSelector->setMinimumWidth(120);
    connect(formatSelector, QOverload<int>::of(&QComboBox::currentIndexChanged), this, &MainWindow::onFormatChanged);
    controlLayout->addWidget(formatSelector);

    convertButton = new QPushButton("Convert All", this);
    convertButton->setStyleSheet("QPushButton { background-color: #4CAF50; color: white; font-weight: bold; padding: 8px 20px; }"
                                 "QPushButton:disabled { background-color: #cccccc; color: #666666; }");
    connect(convertButton, &QPushButton::clicked, this, &MainWindow::onConvertClicked);
    controlLayout->addWidget(convertButton);

    cancelButton = new QPushButton("Cancel", this);
    cancelButton->setStyleSheet("QPushButton { background-color: #f44336; color: white; font-weight: bold; padding: 8px 20px; }");
    cancelButton->setVisible(false);
    connect(cancelButton, &QPushButton::clicked, this, &MainWindow::onCancelClicked);
    controlLayout->addWidget(cancelButton);

    mainLayout->addLayout(controlLayout);

    // Output directory selection
    QHBoxLayout *outputLayout = new QHBoxLayout();
    QLabel *outputLabel = new QLabel("Output folder:", this);
    outputLayout->addWidget(outputLabel);
    
    outputDirLabel = new QLabel(this);
    outputDirLabel->setStyleSheet("QLabel { color: #555; }");
    outputDirLabel->setText(QDir::homePath() + "/Documents/FileConverter_Output");
    outputDirLabel->setWordWrap(true);
    outputLayout->addWidget(outputDirLabel, 1);
    
    browseOutputButton = new QPushButton("Browse...", this);
    connect(browseOutputButton, &QPushButton::clicked, this, [this]() {
        QString dir = QFileDialog::getExistingDirectory(this, "Select Output Directory", outputDirectory);
        if (!dir.isEmpty()) {
            outputDirectory = dir;
            outputDirLabel->setText(dir);
        }
    });
    outputLayout->addWidget(browseOutputButton);
    
    mainLayout->addLayout(outputLayout);

    // Progress section
    QHBoxLayout *progressLayout = new QHBoxLayout();
    
    statusLabel = new QLabel("Ready", this);
    progressLayout->addWidget(statusLabel);

    progressBar = new QProgressBar(this);
    progressBar->setVisible(false);
    progressLayout->addWidget(progressBar, 1);
    
    timeLabel = new QLabel("", this);
    timeLabel->setStyleSheet("QLabel { color: #666; font-size: 11px; }");
    timeLabel->setVisible(false);
    progressLayout->addWidget(timeLabel);

    mainLayout->addLayout(progressLayout);

    // Status bar
    statusBar()->showMessage("Ready - Drag and drop files or click 'Add Files'");
}

void MainWindow::onFilesDropped(const QStringList &filePaths)
{
    addFilesToList(filePaths);
}

void MainWindow::onAddFilesClicked()
{
    QStringList files = QFileDialog::getOpenFileNames(
        this,
        "Select Files to Convert",
        QString(),
        "All Supported Files (*.docx *.pptx *.pdf *.jpg *.jpeg *.png *.webp *.heic *.heif);;Documents (*.docx *.pptx *.pdf);;Images (*.jpg *.jpeg *.png *.webp *.heic *.heif);;All Files (*.*)"
    );

    if (!files.isEmpty()) {
        addFilesToList(files);
    }
}

void MainWindow::addFilesToList(const QStringList &filePaths)
{
    for (const QString &filePath : filePaths) {
        // Check if file already exists
        if (findFileRow(filePath) != -1) {
            continue;
        }

        QFileInfo fileInfo(filePath);
        int row = fileListTable->rowCount();
        fileListTable->insertRow(row);

        fileListTable->setItem(row, 0, new QTableWidgetItem(fileInfo.fileName()));
        fileListTable->setItem(row, 1, new QTableWidgetItem(filePath));
        
        Converter::FileFormat format = Converter::detectFormat(filePath);
        fileListTable->setItem(row, 2, new QTableWidgetItem(Converter::formatToString(format)));
        fileListTable->setItem(row, 3, new QTableWidgetItem("Pending"));
    }

    statusBar()->showMessage(QString("%1 file(s) ready").arg(fileListTable->rowCount()));
    updateConvertButtonState();
}

int MainWindow::findFileRow(const QString &filePath)
{
    for (int i = 0; i < fileListTable->rowCount(); ++i) {
        if (fileListTable->item(i, 1)->text() == filePath) {
            return i;
        }
    }
    return -1;
}

void MainWindow::onClearClicked()
{
    fileListTable->setRowCount(0);
    statusBar()->showMessage("File list cleared");
    updateConvertButtonState();
}

void MainWindow::onRemoveSelectedClicked()
{
    QList<QTableWidgetItem *> selectedItems = fileListTable->selectedItems();
    QSet<int> rowsToRemove;
    
    for (QTableWidgetItem *item : selectedItems) {
        rowsToRemove.insert(item->row());
    }

    QList<int> sortedRows = rowsToRemove.values();
    std::sort(sortedRows.begin(), sortedRows.end(), std::greater<int>());

    for (int row : sortedRows) {
        fileListTable->removeRow(row);
    }

    statusBar()->showMessage(QString("%1 file(s) removed").arg(sortedRows.count()));
    updateConvertButtonState();
}

void MainWindow::onConvertClicked()
{
    if (fileListTable->rowCount() == 0) {
        QMessageBox::warning(this, "No Files", "Please add files to convert.");
        return;
    }

    // Ask user where to save output
    QString dir = QFileDialog::getExistingDirectory(this, "Select Output Directory", outputDirectory);
    if (dir.isEmpty()) {
        return; // User cancelled
    }
    outputDirectory = dir;
    outputDirLabel->setText(dir);
    
    // Create output directory if it doesn't exist
    QDir().mkpath(outputDirectory);
    
    // Set output directory on converter
    converter->setOutputDirectory(outputDirectory);

    Converter::FileFormat targetFormat = static_cast<Converter::FileFormat>(
        formatSelector->currentData().toInt()
    );

    totalFiles = fileListTable->rowCount();
    processedFiles = 0;
    lastOutputPath = outputDirectory;

    progressBar->setMaximum(totalFiles);
    progressBar->setValue(0);
    progressBar->setVisible(true);
    timeLabel->setVisible(true);
    timeLabel->setText("Estimating...");

    convertButton->setVisible(false);
    cancelButton->setVisible(true);
    addFilesButton->setEnabled(false);
    clearButton->setEnabled(false);
    removeButton->setEnabled(false);
    formatSelector->setEnabled(false);
    browseOutputButton->setEnabled(false);

    // Start elapsed timer
    elapsedTimer.start();
    progressTimer->start(500); // Update every 500ms

    // Queue all files for conversion (converter handles parallel execution)
    for (int i = 0; i < fileListTable->rowCount(); ++i) {
        QString filePath = fileListTable->item(i, 1)->text();
        fileListTable->item(i, 3)->setText("Queued");
        converter->convertFile(filePath, targetFormat);
    }
}

void MainWindow::onConversionStarted(const QString &filePath)
{
    int row = findFileRow(filePath);
    if (row != -1) {
        fileListTable->item(row, 3)->setText("Converting...");
        statusLabel->setText(QString("Converting: %1").arg(QFileInfo(filePath).fileName()));
    }
}

void MainWindow::onConversionFinished(const QString &filePath, Converter::ConversionStatus status, const QString &outputPath)
{
    int row = findFileRow(filePath);
    if (row != -1) {
        switch (status) {
            case Converter::ConversionStatus::Success:
                fileListTable->item(row, 3)->setText("✓ Success");
                if (!outputPath.isEmpty()) {
                    lastOutputPath = QFileInfo(outputPath).absolutePath();
                }
                break;
            case Converter::ConversionStatus::Failed:
                fileListTable->item(row, 3)->setText("✗ Failed");
                break;
            case Converter::ConversionStatus::Unsupported:
                fileListTable->item(row, 3)->setText("⚠ Unsupported");
                break;
            case Converter::ConversionStatus::Cancelled:
                fileListTable->item(row, 3)->setText("⊘ Cancelled");
                break;
        }
    }

    processedFiles++;
    progressBar->setValue(processedFiles);
    
    // Update progress with time estimate
    qint64 elapsed = elapsedTimer.elapsed();
    if (processedFiles > 0) {
        qint64 avgTimePerFile = elapsed / processedFiles;
        qint64 remainingFiles = totalFiles - processedFiles;
        qint64 estimatedRemaining = avgTimePerFile * remainingFiles;
        
        statusLabel->setText(QString("Converting: %1/%2 files").arg(processedFiles).arg(totalFiles));
        timeLabel->setText(QString("Elapsed: %1 | Remaining: ~%2")
                          .arg(formatElapsedTime(elapsed))
                          .arg(formatRemainingTime(estimatedRemaining)));
    }
}

void MainWindow::onCancelClicked()
{
    converter->cancelAll();
    progressTimer->stop();
    
    // Mark remaining queued files as cancelled
    for (int i = 0; i < fileListTable->rowCount(); ++i) {
        QString status = fileListTable->item(i, 3)->text();
        if (status == "Queued" || status == "Converting...") {
            fileListTable->item(i, 3)->setText("⊘ Cancelled");
        }
    }
    
    statusBar()->showMessage("Conversion cancelled");
    
    // Reset UI
    progressBar->setVisible(false);
    timeLabel->setVisible(false);
    convertButton->setVisible(true);
    cancelButton->setVisible(false);
    convertButton->setEnabled(true);
    addFilesButton->setEnabled(true);
    clearButton->setEnabled(true);
    removeButton->setEnabled(true);
    formatSelector->setEnabled(true);
    browseOutputButton->setEnabled(true);
}

void MainWindow::onAllConversionsFinished()
{
    progressTimer->stop();
    qint64 totalTime = elapsedTimer.elapsed();
    
    statusLabel->setText(QString("Completed: %1/%2 files in %3").arg(processedFiles).arg(totalFiles).arg(formatElapsedTime(totalTime)));
    statusBar()->showMessage("Conversion completed");
    progressBar->setVisible(false);
    timeLabel->setVisible(false);
    convertButton->setVisible(true);
    cancelButton->setVisible(false);
    convertButton->setEnabled(true);
    addFilesButton->setEnabled(true);
    clearButton->setEnabled(true);
    removeButton->setEnabled(true);
    formatSelector->setEnabled(true);
    browseOutputButton->setEnabled(true);
    
    // Ask user if they want to open the output folder
    QMessageBox::StandardButton reply = QMessageBox::question(
        this, 
        "Conversion Complete",
        QString("Successfully converted %1 files.\n\nWould you like to open the output folder?").arg(processedFiles),
        QMessageBox::Yes | QMessageBox::No
    );
    
    if (reply == QMessageBox::Yes) {
        QDesktopServices::openUrl(QUrl::fromLocalFile(lastOutputPath));
    }
    
    updateConvertButtonState();
}

void MainWindow::onConversionError(const QString &filePath, const QString &errorMessage)
{
    int row = findFileRow(filePath);
    if (row != -1) {
        fileListTable->item(row, 3)->setText("✗ Error");
    }

    statusBar()->showMessage(QString("Error: %1").arg(errorMessage));
}

void MainWindow::onFormatChanged(int index)
{
    Q_UNUSED(index);
    updateConvertButtonState();
}

void MainWindow::updateConvertButtonState()
{
    if (fileListTable->rowCount() == 0) {
        convertButton->setEnabled(false);
        convertButton->setToolTip("Add files to convert");
        return;
    }
    
    Converter::FileFormat targetFormat = static_cast<Converter::FileFormat>(
        formatSelector->currentData().toInt()
    );
    
    // Check if at least one file can be converted to the target format
    bool hasConvertibleFiles = false;
    int convertibleCount = 0;
    
    for (int i = 0; i < fileListTable->rowCount(); ++i) {
        QString filePath = fileListTable->item(i, 1)->text();
        Converter::FileFormat sourceFormat = Converter::detectFormat(filePath);
        
        if (canConvertToFormat(sourceFormat, targetFormat)) {
            hasConvertibleFiles = true;
            convertibleCount++;
        }
    }
    
    convertButton->setEnabled(hasConvertibleFiles);
    
    if (!hasConvertibleFiles) {
        convertButton->setToolTip("No files can be converted to the selected format");
    } else {
        convertButton->setToolTip(QString("%1 file(s) can be converted to %2")
                                 .arg(convertibleCount)
                                 .arg(Converter::formatToString(targetFormat)));
    }
}

bool MainWindow::canConvertToFormat(Converter::FileFormat sourceFormat, Converter::FileFormat targetFormat)
{
    if (sourceFormat == targetFormat) {
        return false; // Same format, no conversion needed
    }
    
    // Document conversions: DOCX/PPTX <-> PDF
    if ((sourceFormat == Converter::FileFormat::DOCX || sourceFormat == Converter::FileFormat::PPTX) &&
        targetFormat == Converter::FileFormat::PDF) {
        return true;
    }
    if (sourceFormat == Converter::FileFormat::PDF &&
        (targetFormat == Converter::FileFormat::DOCX || targetFormat == Converter::FileFormat::PPTX)) {
        return true;
    }
    
    // Image conversions: JPG/PNG/WEBP/HEIC -> JPG/PNG/WEBP
    // Note: HEIC can be SOURCE but not TARGET (convert FROM heic, not TO heic)
    bool sourceIsImage = (sourceFormat == Converter::FileFormat::JPG ||
                         sourceFormat == Converter::FileFormat::PNG ||
                         sourceFormat == Converter::FileFormat::WEBP ||
                         sourceFormat == Converter::FileFormat::HEIC);
    bool targetIsImage = (targetFormat == Converter::FileFormat::JPG ||
                         targetFormat == Converter::FileFormat::PNG ||
                         targetFormat == Converter::FileFormat::WEBP);
    
    if (sourceIsImage && targetIsImage) {
        return true;
    }
    
    return false;
}

QString MainWindow::formatElapsedTime(qint64 ms)
{
    if (ms < 1000) {
        return QString("%1ms").arg(ms);
    }
    
    qint64 seconds = ms / 1000;
    qint64 minutes = seconds / 60;
    seconds = seconds % 60;
    
    if (minutes > 0) {
        return QString("%1m %2s").arg(minutes).arg(seconds);
    }
    return QString("%1s").arg(seconds);
}

QString MainWindow::formatRemainingTime(qint64 ms)
{
    if (ms < 1000) {
        return "< 1s";
    }
    
    qint64 seconds = ms / 1000;
    qint64 minutes = seconds / 60;
    seconds = seconds % 60;
    
    if (minutes > 0) {
        return QString("%1m %2s").arg(minutes).arg(seconds);
    }
    return QString("%1s").arg(seconds);
}

void MainWindow::updateProgressTimer()
{
    if (!elapsedTimer.isValid() || processedFiles == 0) {
        return;
    }
    
    qint64 elapsed = elapsedTimer.elapsed();
    qint64 avgTimePerFile = elapsed / processedFiles;
    qint64 remainingFiles = totalFiles - processedFiles;
    qint64 estimatedRemaining = avgTimePerFile * remainingFiles;
    
    timeLabel->setText(QString("Elapsed: %1 | Remaining: ~%2")
                      .arg(formatElapsedTime(elapsed))
                      .arg(formatRemainingTime(estimatedRemaining)));
}

void MainWindow::setupMenuBar()
{
    QMenuBar *menuBar = this->menuBar();
    
    // File menu
    QMenu *fileMenu = menuBar->addMenu("&File");
    
    QAction *addFilesAction = fileMenu->addAction("&Add Files...");
    connect(addFilesAction, &QAction::triggered, this, &MainWindow::onAddFilesClicked);
    
    fileMenu->addSeparator();
    
    QAction *exitAction = fileMenu->addAction("E&xit");
    connect(exitAction, &QAction::triggered, this, &QMainWindow::close);
    
    // Integration menu (Windows Shell Integration)
    QMenu *integrationMenu = menuBar->addMenu("&Integration");
    
    // Context Menu submenu
    QMenu *contextMenuSubmenu = integrationMenu->addMenu("Context Menu");
    
    installContextMenuAction = contextMenuSubmenu->addAction("&Install Context Menu");
    installContextMenuAction->setToolTip("Adds 'Convert with FileConverter' to right-click menu for supported file types");
    connect(installContextMenuAction, &QAction::triggered, this, &MainWindow::onInstallContextMenu);
    
    removeContextMenuAction = contextMenuSubmenu->addAction("&Remove Context Menu");
    removeContextMenuAction->setToolTip("Removes the context menu entries (clean uninstall)");
    connect(removeContextMenuAction, &QAction::triggered, this, &MainWindow::onRemoveContextMenu);
    
    integrationMenu->addSeparator();
    
    // Send To submenu
    QMenu *sendToSubmenu = integrationMenu->addMenu("Send To Folder");
    
    installSendToAction = sendToSubmenu->addAction("&Add to Send To");
    installSendToAction->setToolTip("Adds FileConverter to the 'Send To' right-click menu (safer alternative)");
    connect(installSendToAction, &QAction::triggered, this, &MainWindow::onInstallSendTo);
    
    removeSendToAction = sendToSubmenu->addAction("&Remove from Send To");
    removeSendToAction->setToolTip("Removes FileConverter from the 'Send To' menu");
    connect(removeSendToAction, &QAction::triggered, this, &MainWindow::onRemoveSendTo);
    
    // Help menu
    QMenu *helpMenu = menuBar->addMenu("&Help");
    
    QAction *aboutAction = helpMenu->addAction("&About");
    connect(aboutAction, &QAction::triggered, this, [this]() {
        QMessageBox::about(this, "About FileConverter",
            "<h3>FileConverter</h3>"
            "<p>Offline file converter for documents and images.</p>"
            "<p><b>Supported conversions:</b></p>"
            "<ul>"
            "<li>DOCX/PPTX ↔ PDF (requires LibreOffice)</li>"
            "<li>JPG, PNG, WEBP, HEIC (requires ImageMagick)</li>"
            "</ul>"
            "<p>Version 1.0</p>");
    });
    
    // Update menu state based on current installation status
    updateIntegrationMenuState();
}

void MainWindow::updateIntegrationMenuState()
{
    bool shellInstalled = ContextMenu::isShellRegistered();
    bool sendToInstalled = ContextMenu::isSendToInstalled();
    
    installContextMenuAction->setEnabled(!shellInstalled);
    removeContextMenuAction->setEnabled(shellInstalled);
    
    installSendToAction->setEnabled(!sendToInstalled);
    removeSendToAction->setEnabled(sendToInstalled);
}

void MainWindow::onInstallContextMenu()
{
    if (ContextMenu::registerShellExtension()) {
        QMessageBox::information(this, "Success", 
            "Context menu installed successfully!\n\n"
            "Right-click on supported files (DOCX, PDF, images) to see 'Convert with FileConverter'.\n\n"
            "Note: On Windows 11, you may need to click 'Show more options' first.");
        updateIntegrationMenuState();
    } else {
        QMessageBox::warning(this, "Error", 
            "Failed to install context menu.\n\n"
            "Try running the application as Administrator.");
    }
}

void MainWindow::onRemoveContextMenu()
{
    if (ContextMenu::unregisterShellExtension()) {
        QMessageBox::information(this, "Success", 
            "Context menu removed successfully!\n\n"
            "All registry entries have been cleaned up.");
        updateIntegrationMenuState();
    } else {
        QMessageBox::warning(this, "Error", 
            "Failed to remove context menu.\n\n"
            "Try running the application as Administrator.");
    }
}

void MainWindow::onInstallSendTo()
{
    if (ContextMenu::createSendToShortcut()) {
        QMessageBox::information(this, "Success", 
            "Send To shortcut created!\n\n"
            "Right-click any file → Send To → FileConverter");
        updateIntegrationMenuState();
    } else {
        QMessageBox::warning(this, "Error", 
            "Failed to create Send To shortcut.");
    }
}

void MainWindow::onRemoveSendTo()
{
    if (ContextMenu::removeSendToShortcut()) {
        QMessageBox::information(this, "Success", 
            "Send To shortcut removed!");
        updateIntegrationMenuState();
    } else {
        QMessageBox::warning(this, "Error", 
            "Failed to remove Send To shortcut.");
    }
}
