#ifndef IMAGECODEC_H
#define IMAGECODEC_H

#include <QImage>
#include <QString>
#include <QByteArray>

// In-process image decode/encode helpers used to avoid shelling out to
// ImageMagick for the common formats. Qt's own QImage already handles
// JPEG/PNG; WEBP (when libwebp is available at build time) and HEIC
// (native macOS decode only) are implemented here.
namespace ImageCodec {

#ifdef HAVE_WEBP
bool decodeWebP(const QByteArray &data, QImage &out);
bool encodeWebP(const QImage &image, QByteArray &out, float quality = 90.0f);
#endif

#ifdef Q_OS_MACOS
bool decodeHeicMac(const QString &path, QImage &out);
#endif

}

#endif // IMAGECODEC_H
