#include "Converter.h"
#include <QFileInfo>
#include <QDir>
#include <QStandardPaths>
#include <QDebug>
#include <QTimer>

Converter::Converter(QObject *parent)
    : QObject(parent), maxParallelConversions(1)  // Use 1 to avoid LibreOffice conflicts
{
    libreOfficePath = findLibreOffice();
    imageMagickPath = findImageMagick();
}

Converter::~Converter()
{
    cancelAll();
}

void Converter::setLibreOfficePath(const QString &path)
{
    libreOfficePath = path;
}

void Converter::setImageMagickPath(const QString &path)
{
    imageMagickPath = path;
}

void Converter::setMaxParallelConversions(int max)
{
    maxParallelConversions = qMax(1, max);
}

void Converter::setOutputDirectory(const QString &path)
{
    outputDirectory = path;
}

bool Converter::isConverting() const
{
    return !activeJobs.isEmpty() || !conversionQueue.isEmpty();
}

int Converter::activeConversions() const
{
    return activeJobs.size();
}

Converter::FileFormat Converter::detectFormat(const QString &filePath)
{
    QString suffix = QFileInfo(filePath).suffix().toLower();
    if (suffix == "docx") return FileFormat::DOCX;
    if (suffix == "pptx") return FileFormat::PPTX;
    if (suffix == "pdf") return FileFormat::PDF;
    if (suffix == "jpg" || suffix == "jpeg") return FileFormat::JPG;
    if (suffix == "png") return FileFormat::PNG;
    if (suffix == "webp") return FileFormat::WEBP;
    if (suffix == "heic" || suffix == "heif") return FileFormat::HEIC;
    return FileFormat::Unknown;
}

QString Converter::formatToString(FileFormat format)
{
    switch (format) {
        case FileFormat::DOCX: return "DOCX";
        case FileFormat::PPTX: return "PPTX";
        case FileFormat::PDF: return "PDF";
        case FileFormat::JPG: return "JPG";
        case FileFormat::PNG: return "PNG";
        case FileFormat::WEBP: return "WEBP";
        case FileFormat::HEIC: return "HEIC";
        default: return "Unknown";
    }
}

QString Converter::formatToExtension(FileFormat format)
{
    switch (format) {
        case FileFormat::DOCX: return "docx";
        case FileFormat::PPTX: return "pptx";
        case FileFormat::PDF: return "pdf";
        case FileFormat::JPG: return "jpg";
        case FileFormat::PNG: return "png";
        case FileFormat::WEBP: return "webp";
        case FileFormat::HEIC: return "heic";
        default: return "";
    }
}

void Converter::convertFile(const QString &inputPath, FileFormat targetFormat)
{
    if (!QFileInfo::exists(inputPath)) {
        emit conversionError(inputPath, "File does not exist");
        return;
    }

    FileFormat sourceFormat = detectFormat(inputPath);
    if (sourceFormat == FileFormat::Unknown) {
        emit conversionError(inputPath, "Unsupported file format");
        return;
    }

    // Check if already converting this file
    if (activeJobs.contains(inputPath)) {
        emit conversionError(inputPath, "File is already being converted");
        return;
    }

    // Add to queue
    conversionQueue.append(qMakePair(inputPath, targetFormat));
    
    // Start conversion if we have capacity
    startNextQueuedConversion();
}

void Converter::startNextQueuedConversion()
{
    while (activeJobs.size() < maxParallelConversions && !conversionQueue.isEmpty()) {
        auto job = conversionQueue.takeFirst();
        QString inputPath = job.first;
        FileFormat targetFormat = job.second;
        
        FileFormat sourceFormat = detectFormat(inputPath);
        QFileInfo fileInfo(inputPath);
        
        // Use outputDirectory if set, otherwise use same directory as input
        QString outDir = outputDirectory.isEmpty() ? fileInfo.absolutePath() : outputDirectory;
        QString outputPath = outDir + "/" + fileInfo.baseName() + "." + formatToExtension(targetFormat);

        emit conversionStarted(inputPath);

        // Document to PDF conversions (DOCX/PPTX -> PDF)
        if ((sourceFormat == FileFormat::DOCX || sourceFormat == FileFormat::PPTX) && targetFormat == FileFormat::PDF) {
            convertDocumentToPDF(inputPath, outputPath);
        }
        // PDF to document conversions (PDF -> DOCX/PPTX)
        else if (sourceFormat == FileFormat::PDF && (targetFormat == FileFormat::DOCX || targetFormat == FileFormat::PPTX)) {
            convertPDFtoDocument(inputPath, outputPath, targetFormat);
        }
        // Image conversions (including HEIC as source - HEIC can be converted TO other formats but not FROM)
        else if ((sourceFormat == FileFormat::JPG || sourceFormat == FileFormat::PNG || sourceFormat == FileFormat::WEBP || sourceFormat == FileFormat::HEIC) &&
                 (targetFormat == FileFormat::JPG || targetFormat == FileFormat::PNG || targetFormat == FileFormat::WEBP)) {
            convertImage(inputPath, outputPath, targetFormat);
        }
        else {
            emit conversionFinished(inputPath, ConversionStatus::Unsupported, "");
        }
    }
}

void Converter::cancelConversion(const QString &inputPath)
{
    // Check queue first
    for (int i = 0; i < conversionQueue.size(); ++i) {
        if (conversionQueue[i].first == inputPath) {
            conversionQueue.removeAt(i);
            emit conversionFinished(inputPath, ConversionStatus::Cancelled, "");
            return;
        }
    }
    
    // Check active jobs
    if (activeJobs.contains(inputPath)) {
        ConversionJob &job = activeJobs[inputPath];
        job.cancelled = true;
        if (job.process) {
            job.process->kill();
        }
    }
}

void Converter::cancelAll()
{
    // Clear queue
    QList<QPair<QString, FileFormat>> queueCopy = conversionQueue;
    conversionQueue.clear();
    
    for (const auto &job : queueCopy) {
        emit conversionFinished(job.first, ConversionStatus::Cancelled, "");
    }
    
    // Kill active processes
    for (auto it = activeJobs.begin(); it != activeJobs.end(); ++it) {
        it.value().cancelled = true;
        if (it.value().process) {
            it.value().process->kill();
        }
    }
}

void Converter::convertDocumentToPDF(const QString &inputPath, const QString &outputPath)
{
    if (libreOfficePath.isEmpty()) {
        emit conversionError(inputPath, "LibreOffice not found. Please install LibreOffice.");
        return;
    }

    QProcess *process = new QProcess(this);
    
    ConversionJob job;
    job.process = process;
    job.inputPath = inputPath;
    job.outputPath = outputPath;
    job.cancelled = false;
    activeJobs[inputPath] = job;
    
    connect(process, QOverload<int, QProcess::ExitStatus>::of(&QProcess::finished),
            this, &Converter::onProcessFinished);
    connect(process, &QProcess::errorOccurred, this, &Converter::onProcessError);

    QFileInfo outputInfo(outputPath);
    QStringList args;
    args << "--headless"
         << "--convert-to" << "pdf"
         << "--outdir" << outputInfo.absolutePath()
         << inputPath;

    process->start(libreOfficePath, args);
}

void Converter::convertPDFtoDocument(const QString &inputPath, const QString &outputPath, FileFormat targetFormat)
{
    if (libreOfficePath.isEmpty()) {
        emit conversionError(inputPath, "LibreOffice not found. Please install LibreOffice.");
        return;
    }

    QProcess *process = new QProcess(this);
    
    ConversionJob job;
    job.process = process;
    job.inputPath = inputPath;
    job.outputPath = outputPath;
    job.cancelled = false;
    activeJobs[inputPath] = job;
    
    connect(process, QOverload<int, QProcess::ExitStatus>::of(&QProcess::finished),
            this, &Converter::onProcessFinished);
    connect(process, &QProcess::errorOccurred, this, &Converter::onProcessError);

    QFileInfo outputInfo(outputPath);
    
    // LibreOffice PDF to DOCX: use writer_pdf_import filter and appropriate export filter
    QString formatStr;
    QString infilter;
    
    if (targetFormat == FileFormat::PPTX) {
        // PDF to PPTX - use Draw's PDF import
        infilter = "draw_pdf_import";
        formatStr = "pptx";
    } else {
        // PDF to DOCX - use Writer's PDF import
        infilter = "writer_pdf_import";
        formatStr = "docx";
    }
    
    QStringList args;
    args << "--headless"
         << QString("--infilter=%1").arg(infilter)
         << "--convert-to" << formatStr
         << "--outdir" << outputInfo.absolutePath()
         << inputPath;

    process->start(libreOfficePath, args);
}

void Converter::convertImage(const QString &inputPath, const QString &outputPath, FileFormat targetFormat)
{
    Q_UNUSED(targetFormat);
    
    if (imageMagickPath.isEmpty()) {
        emit conversionError(inputPath, "ImageMagick not found. Please install ImageMagick.");
        return;
    }

    QProcess *process = new QProcess(this);
    
    ConversionJob job;
    job.process = process;
    job.inputPath = inputPath;
    job.outputPath = outputPath;
    job.cancelled = false;
    activeJobs[inputPath] = job;
    
    connect(process, QOverload<int, QProcess::ExitStatus>::of(&QProcess::finished),
            this, &Converter::onProcessFinished);
    connect(process, &QProcess::errorOccurred, this, &Converter::onProcessError);

    QStringList args;
    args << inputPath
         << outputPath;

    process->start(imageMagickPath, args);
}

void Converter::onProcessFinished(int exitCode, QProcess::ExitStatus exitStatus)
{
    QProcess *process = qobject_cast<QProcess*>(sender());
    if (!process) return;
    
    // Find the job for this process
    QString inputPath;
    QString outputPath;
    bool cancelled = false;
    
    for (auto it = activeJobs.begin(); it != activeJobs.end(); ++it) {
        if (it.value().process == process) {
            inputPath = it.value().inputPath;
            outputPath = it.value().outputPath;
            cancelled = it.value().cancelled;
            activeJobs.erase(it);
            break;
        }
    }
    
    process->deleteLater();
    
    if (inputPath.isEmpty()) return;
    
    if (cancelled) {
        emit conversionFinished(inputPath, ConversionStatus::Cancelled, "");
        finalizeConversion();
    } else if (exitStatus == QProcess::NormalExit && exitCode == 0) {
        // Schedule non-blocking file existence check
        scheduleFileCheck(inputPath, outputPath);
    } else {
        QString errorOutput = process->readAllStandardError();
        QString stdOutput = process->readAllStandardOutput();
        QString fullError = errorOutput.isEmpty() ? stdOutput : errorOutput;
        if (fullError.isEmpty()) {
            fullError = QString("Process exited with code %1").arg(exitCode);
        }
        emit conversionError(inputPath, "Conversion failed: " + fullError);
        finalizeConversion();
    }
}

void Converter::scheduleFileCheck(const QString &inputPath, const QString &outputPath)
{
    PendingFileCheck check;
    check.inputPath = inputPath;
    check.outputPath = outputPath;
    check.retryCount = 0;
    check.timer = new QTimer(this);
    check.timer->setSingleShot(true);
    
    pendingFileChecks[inputPath] = check;
    
    connect(check.timer, &QTimer::timeout, this, &Converter::checkFileExists);
    check.timer->start(50); // First check after 50ms
}

void Converter::checkFileExists()
{
    QTimer *timer = qobject_cast<QTimer*>(sender());
    if (!timer) return;
    
    // Find which pending check this timer belongs to
    QString inputPath;
    for (auto it = pendingFileChecks.begin(); it != pendingFileChecks.end(); ++it) {
        if (it.value().timer == timer) {
            inputPath = it.key();
            break;
        }
    }
    
    if (inputPath.isEmpty()) {
        timer->deleteLater();
        return;
    }
    
    PendingFileCheck &check = pendingFileChecks[inputPath];
    QString outputPath = check.outputPath;
    
    // Check if file exists and has content
    if (QFileInfo::exists(outputPath)) {
        QFileInfo fi(outputPath);
        if (fi.size() > 0) {
            // Success!
            timer->deleteLater();
            pendingFileChecks.remove(inputPath);
            emit conversionFinished(inputPath, ConversionStatus::Success, outputPath);
            finalizeConversion();
            return;
        }
    }
    
    check.retryCount++;
    
    if (check.retryCount < 20) { // Max 20 retries = 2 seconds total
        timer->start(100); // Retry after 100ms
    } else {
        // Final attempt: search for file with matching pattern
        QFileInfo inputInfo(inputPath);
        QFileInfo outputInfo(outputPath);
        QString outDir = outputInfo.absolutePath();
        QString baseName = inputInfo.baseName();
        QString ext = outputInfo.suffix();
        
        QDir dir(outDir);
        QStringList filters;
        filters << baseName + "." + ext;
        filters << baseName + "*." + ext;
        QStringList matchingFiles = dir.entryList(filters, QDir::Files, QDir::Time);
        
        timer->deleteLater();
        pendingFileChecks.remove(inputPath);
        
        if (!matchingFiles.isEmpty()) {
            QString foundPath = outDir + "/" + matchingFiles.first();
            emit conversionFinished(inputPath, ConversionStatus::Success, foundPath);
        } else {
            emit conversionError(inputPath, "Output file was not created. Check if LibreOffice/ImageMagick is installed correctly.");
        }
        finalizeConversion();
    }
}

void Converter::finalizeConversion()
{
    // Start next queued conversion
    startNextQueuedConversion();
    
    // Check if all done (no active jobs, no queue, no pending file checks)
    if (activeJobs.isEmpty() && conversionQueue.isEmpty() && pendingFileChecks.isEmpty()) {
        emit allConversionsFinished();
    }
}

void Converter::onProcessError(QProcess::ProcessError error)
{
    QProcess *process = qobject_cast<QProcess*>(sender());
    if (!process) return;
    
    // Find the job for this process
    QString inputPath;
    
    for (auto it = activeJobs.begin(); it != activeJobs.end(); ++it) {
        if (it.value().process == process) {
            inputPath = it.value().inputPath;
            activeJobs.erase(it);
            break;
        }
    }
    
    process->deleteLater();
    
    if (inputPath.isEmpty()) return;
    
    QString errorMsg;
    switch (error) {
        case QProcess::FailedToStart:
            errorMsg = "Failed to start conversion tool";
            break;
        case QProcess::Crashed:
            errorMsg = "Conversion tool crashed";
            break;
        default:
            errorMsg = "Unknown error occurred";
            break;
    }
    emit conversionError(inputPath, errorMsg);
    
    // Start next queued conversion
    startNextQueuedConversion();
    
    // Check if all done
    if (activeJobs.isEmpty() && conversionQueue.isEmpty()) {
        emit allConversionsFinished();
    }
}

QString Converter::findLibreOffice()
{
    // Common LibreOffice installation paths on Windows
    QStringList possiblePaths = {
        "C:/Program Files/LibreOffice/program/soffice.exe",
        "C:/Program Files (x86)/LibreOffice/program/soffice.exe",
        QDir::homePath() + "/AppData/Local/Programs/LibreOffice/program/soffice.exe"
    };

    for (const QString &path : possiblePaths) {
        if (QFileInfo::exists(path)) {
            return path;
        }
    }

    // Try to find in PATH
    QString pathEnv = qEnvironmentVariable("PATH");
    QStringList pathDirs = pathEnv.split(';', Qt::SkipEmptyParts);
    for (const QString &dir : pathDirs) {
        QString sofficePath = dir + "/soffice.exe";
        if (QFileInfo::exists(sofficePath)) {
            return sofficePath;
        }
    }

    return QString();
}

QString Converter::findImageMagick()
{
    // Common ImageMagick installation paths on Windows
    QStringList possiblePaths;
    
    // Check Program Files directories
    QDir programFiles("C:/Program Files");
    QStringList imageMagickDirs = programFiles.entryList(QStringList() << "ImageMagick*", QDir::Dirs);
    for (const QString &dir : imageMagickDirs) {
        possiblePaths << "C:/Program Files/" + dir + "/magick.exe";
    }

    QDir programFilesX86("C:/Program Files (x86)");
    imageMagickDirs = programFilesX86.entryList(QStringList() << "ImageMagick*", QDir::Dirs);
    for (const QString &dir : imageMagickDirs) {
        possiblePaths << "C:/Program Files (x86)/" + dir + "/magick.exe";
    }

    for (const QString &path : possiblePaths) {
        if (QFileInfo::exists(path)) {
            return path;
        }
    }

    // Try to find in PATH
    QString pathEnv = qEnvironmentVariable("PATH");
    QStringList pathDirs = pathEnv.split(';', Qt::SkipEmptyParts);
    for (const QString &dir : pathDirs) {
        QString magickPath = dir + "/magick.exe";
        if (QFileInfo::exists(magickPath)) {
            return magickPath;
        }
    }

    return QString();
}
