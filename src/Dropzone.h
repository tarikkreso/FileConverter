#ifndef DROPZONE_H
#define DROPZONE_H

#include <QWidget>
#include <QStringList>

class Dropzone : public QWidget
{
    Q_OBJECT

public:
    explicit Dropzone(QWidget *parent = nullptr);

signals:
    void filesDropped(const QStringList &filePaths);

protected:
    void dragEnterEvent(QDragEnterEvent *event) override;
    void dragMoveEvent(QDragMoveEvent *event) override;
    void dropEvent(QDropEvent *event) override;
    void paintEvent(QPaintEvent *event) override;

private:
    bool isDragging;
};

#endif // DROPZONE_H
