#ifndef PREFERENCESDIALOG_H
#define PREFERENCESDIALOG_H

#include <QDialog>

class QLineEdit;
class QCheckBox;
class QSlider;
class QLabel;
class QSpinBox;

class PreferencesDialog : public QDialog
{
    Q_OBJECT

public:
    explicit PreferencesDialog(QWidget *parent = nullptr);

    void setOutputDirectory(const QString &dir);
    QString outputDirectory() const;

    void setRememberOutputDirectory(bool remember);
    bool rememberOutputDirectory() const;

    void setJpgQuality(int quality);
    int jpgQuality() const;

    void setMaxParallelConversions(int count);
    int maxParallelConversions() const;

private slots:
    void onBrowseOutputDirectory();

private:
    QLineEdit *outputDirEdit;
    QCheckBox *rememberFolderCheck;
    QSlider *jpgQualitySlider;
    QLabel *jpgQualityValueLabel;
    QSpinBox *maxParallelSpin;
};

#endif // PREFERENCESDIALOG_H
