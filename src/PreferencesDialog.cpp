#include "PreferencesDialog.h"
#include <QVBoxLayout>
#include <QHBoxLayout>
#include <QFormLayout>
#include <QLineEdit>
#include <QCheckBox>
#include <QSlider>
#include <QLabel>
#include <QSpinBox>
#include <QPushButton>
#include <QDialogButtonBox>
#include <QFileDialog>

PreferencesDialog::PreferencesDialog(QWidget *parent)
    : QDialog(parent)
{
    setWindowTitle("Preferences");
    setMinimumWidth(440);

    QVBoxLayout *mainLayout = new QVBoxLayout(this);
    QFormLayout *form = new QFormLayout();

    QHBoxLayout *dirLayout = new QHBoxLayout();
    outputDirEdit = new QLineEdit(this);
    dirLayout->addWidget(outputDirEdit, 1);
    QPushButton *browseButton = new QPushButton("Browse...", this);
    connect(browseButton, &QPushButton::clicked, this, &PreferencesDialog::onBrowseOutputDirectory);
    dirLayout->addWidget(browseButton);
    form->addRow("Default output folder:", dirLayout);

    rememberFolderCheck = new QCheckBox("Remember last used output folder between launches", this);
    form->addRow(QString(), rememberFolderCheck);

    QHBoxLayout *qualityLayout = new QHBoxLayout();
    jpgQualitySlider = new QSlider(Qt::Horizontal, this);
    jpgQualitySlider->setRange(1, 100);
    jpgQualityValueLabel = new QLabel(this);
    jpgQualityValueLabel->setMinimumWidth(28);
    connect(jpgQualitySlider, &QSlider::valueChanged, this, [this](int value) {
        jpgQualityValueLabel->setText(QString::number(value));
    });
    qualityLayout->addWidget(jpgQualitySlider, 1);
    qualityLayout->addWidget(jpgQualityValueLabel);
    form->addRow("JPG quality:", qualityLayout);

    maxParallelSpin = new QSpinBox(this);
    maxParallelSpin->setRange(1, 4);
    maxParallelSpin->setToolTip(
        "Increasing this may cause instability: LibreOffice does not reliably handle "
        "multiple concurrent headless document conversions.");
    form->addRow("Max parallel conversions:", maxParallelSpin);

    mainLayout->addLayout(form);

    QDialogButtonBox *buttonBox = new QDialogButtonBox(QDialogButtonBox::Ok | QDialogButtonBox::Cancel, this);
    connect(buttonBox, &QDialogButtonBox::accepted, this, &QDialog::accept);
    connect(buttonBox, &QDialogButtonBox::rejected, this, &QDialog::reject);
    mainLayout->addWidget(buttonBox);
}

void PreferencesDialog::setOutputDirectory(const QString &dir)
{
    outputDirEdit->setText(dir);
}

QString PreferencesDialog::outputDirectory() const
{
    return outputDirEdit->text();
}

void PreferencesDialog::setRememberOutputDirectory(bool remember)
{
    rememberFolderCheck->setChecked(remember);
}

bool PreferencesDialog::rememberOutputDirectory() const
{
    return rememberFolderCheck->isChecked();
}

void PreferencesDialog::setJpgQuality(int quality)
{
    jpgQualitySlider->setValue(quality);
    jpgQualityValueLabel->setText(QString::number(quality));
}

int PreferencesDialog::jpgQuality() const
{
    return jpgQualitySlider->value();
}

void PreferencesDialog::setMaxParallelConversions(int count)
{
    maxParallelSpin->setValue(count);
}

int PreferencesDialog::maxParallelConversions() const
{
    return maxParallelSpin->value();
}

void PreferencesDialog::onBrowseOutputDirectory()
{
    QString dir = QFileDialog::getExistingDirectory(this, "Select Default Output Directory", outputDirEdit->text());
    if (!dir.isEmpty()) {
        outputDirEdit->setText(dir);
    }
}
