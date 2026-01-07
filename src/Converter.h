#ifndef CONVERTER_H
#define CONVERTER_H

#include <QObject>
#include <QString>
#include <QProcess>
#include <QMap>

class Converter : public QObject
{
    Q_OBJECT

public:
    enum class ConversionStatus {
        Success,
        Failed,
        Unsupported,
        Cancelled
    };

    enum class FileFormat {
        DOCX,
        PPTX,
        PDF,
        JPG,
        PNG,
        WEBP,
        HEIC,
        Unknown
    };

    explicit Converter(QObject *parent = nullptr);
    ~Converter();

    void convertFile(const QString &inputPath, FileFormat targetFormat);
    void cancelConversion(const QString &inputPath);
    void cancelAll();
    bool isConverting() const;
    int activeConversions() const;
    
    static FileFormat detectFormat(const QString &filePath);
    static QString formatToString(FileFormat format);
    static QString formatToExtension(FileFormat format);

    void setLibreOfficePath(const QString &path);
    void setImageMagickPath(const QString &path);
    void setMaxParallelConversions(int max);
    void setOutputDirectory(const QString &path);

signals:
    void conversionStarted(const QString &filePath);
    void conversionProgress(const QString &filePath, int percent);
    void conversionFinished(const QString &filePath, ConversionStatus status, const QString &outputPath);
    void conversionError(const QString &filePath, const QString &errorMessage);
    void allConversionsFinished();

private slots:
    void onProcessFinished(int exitCode, QProcess::ExitStatus exitStatus);
    void onProcessError(QProcess::ProcessError error);
    void checkFileExists();

private:
    struct ConversionJob {
        QProcess *process;
        QString inputPath;
        QString outputPath;
        bool cancelled;
    };
    
    struct PendingFileCheck {
        QString inputPath;
        QString outputPath;
        int retryCount;
        QTimer *timer;
    };

    void convertDocumentToPDF(const QString &inputPath, const QString &outputPath);
    void convertPDFtoDocument(const QString &inputPath, const QString &outputPath, FileFormat targetFormat);
    void convertImage(const QString &inputPath, const QString &outputPath, FileFormat targetFormat);
    void startNextQueuedConversion();
    void scheduleFileCheck(const QString &inputPath, const QString &outputPath);
    void finalizeConversion();
    QString findLibreOffice();
    QString findImageMagick();

    QString libreOfficePath;
    QString imageMagickPath;
    QString outputDirectory;
    
    // Active conversions: key = inputPath
    QMap<QString, ConversionJob> activeJobs;
    
    // Pending file existence checks
    QMap<QString, PendingFileCheck> pendingFileChecks;
    
    // Queue for pending conversions
    QList<QPair<QString, FileFormat>> conversionQueue;
    
    int maxParallelConversions;
};

#endif // CONVERTER_H
