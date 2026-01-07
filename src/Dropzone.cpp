#include "Dropzone.h"
#include <QDragEnterEvent>
#include <QDropEvent>
#include <QMimeData>
#include <QPainter>
#include <QUrl>

Dropzone::Dropzone(QWidget *parent)
    : QWidget(parent), isDragging(false)
{
    setAcceptDrops(true);
    setMinimumSize(400, 150);
}

void Dropzone::dragEnterEvent(QDragEnterEvent *event)
{
    if (event->mimeData()->hasUrls()) {
        event->acceptProposedAction();
        isDragging = true;
        update();
    }
}

void Dropzone::dragMoveEvent(QDragMoveEvent *event)
{
    if (event->mimeData()->hasUrls()) {
        event->acceptProposedAction();
    }
}

void Dropzone::dropEvent(QDropEvent *event)
{
    isDragging = false;
    update();

    const QMimeData *mimeData = event->mimeData();
    if (mimeData->hasUrls()) {
        QStringList filePaths;
        QList<QUrl> urls = mimeData->urls();
        for (const QUrl &url : urls) {
            if (url.isLocalFile()) {
                filePaths.append(url.toLocalFile());
            }
        }
        if (!filePaths.isEmpty()) {
            emit filesDropped(filePaths);
        }
        event->acceptProposedAction();
    }
}

void Dropzone::paintEvent(QPaintEvent *event)
{
    Q_UNUSED(event);
    QPainter painter(this);
    painter.setRenderHint(QPainter::Antialiasing);

    // Draw background
    if (isDragging) {
        painter.setBrush(QColor(230, 240, 255));
        painter.setPen(QPen(QColor(100, 150, 255), 2, Qt::DashLine));
    } else {
        painter.setBrush(QColor(245, 245, 245));
        painter.setPen(QPen(QColor(200, 200, 200), 2, Qt::DashLine));
    }
    painter.drawRoundedRect(rect().adjusted(5, 5, -5, -5), 10, 10);

    // Draw text
    painter.setPen(QColor(100, 100, 100));
    QFont font = painter.font();
    font.setPointSize(12);
    painter.setFont(font);
    painter.drawText(rect(), Qt::AlignCenter, "Drag & Drop Files Here\n\nOr use the buttons below");
}
