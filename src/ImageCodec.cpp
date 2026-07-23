#include "ImageCodec.h"

#ifdef HAVE_WEBP
#include <webp/decode.h>
#include <webp/encode.h>
#endif

#ifdef Q_OS_MACOS
#include <ImageIO/ImageIO.h>
#include <CoreGraphics/CoreGraphics.h>
#include <CoreFoundation/CoreFoundation.h>
#endif

namespace ImageCodec {

#ifdef HAVE_WEBP
bool decodeWebP(const QByteArray &data, QImage &out)
{
    int width = 0;
    int height = 0;
    uint8_t *rgba = WebPDecodeRGBA(reinterpret_cast<const uint8_t *>(data.constData()), data.size(), &width, &height);
    if (!rgba) {
        return false;
    }

    QImage image(rgba, width, height, width * 4, QImage::Format_RGBA8888);
    out = image.copy(); // deep copy before freeing the libwebp buffer
    WebPFree(rgba);
    return !out.isNull();
}

bool encodeWebP(const QImage &imageIn, QByteArray &out, float quality)
{
    QImage image = imageIn.convertToFormat(QImage::Format_RGBA8888);
    if (image.isNull()) {
        return false;
    }

    uint8_t *encoded = nullptr;
    size_t size = WebPEncodeRGBA(image.constBits(), image.width(), image.height(),
                                  static_cast<int>(image.bytesPerLine()), quality, &encoded);
    if (size == 0 || !encoded) {
        return false;
    }

    out = QByteArray(reinterpret_cast<const char *>(encoded), static_cast<int>(size));
    WebPFree(encoded);
    return true;
}
#endif

#ifdef Q_OS_MACOS
bool decodeHeicMac(const QString &path, QImage &out)
{
    CFStringRef cfPath = CFStringCreateWithCString(kCFAllocatorDefault, path.toUtf8().constData(), kCFStringEncodingUTF8);
    if (!cfPath) {
        return false;
    }
    CFURLRef url = CFURLCreateWithFileSystemPath(kCFAllocatorDefault, cfPath, kCFURLPOSIXPathStyle, false);
    CFRelease(cfPath);
    if (!url) {
        return false;
    }

    CGImageSourceRef source = CGImageSourceCreateWithURL(url, nullptr);
    CFRelease(url);
    if (!source) {
        return false;
    }

    CGImageRef cgImage = CGImageSourceCreateImageAtIndex(source, 0, nullptr);
    CFRelease(source);
    if (!cgImage) {
        return false;
    }

    const size_t width = CGImageGetWidth(cgImage);
    const size_t height = CGImageGetHeight(cgImage);

    QImage image(static_cast<int>(width), static_cast<int>(height), QImage::Format_RGBA8888);
    image.fill(Qt::transparent);

    CGColorSpaceRef colorSpace = CGColorSpaceCreateDeviceRGB();
    CGContextRef context = CGBitmapContextCreate(
        image.bits(), width, height, 8, static_cast<size_t>(image.bytesPerLine()),
        colorSpace, kCGImageAlphaPremultipliedLast | kCGBitmapByteOrder32Big);
    CGColorSpaceRelease(colorSpace);

    if (!context) {
        CGImageRelease(cgImage);
        return false;
    }

    CGContextDrawImage(context, CGRectMake(0, 0, width, height), cgImage);
    CGContextRelease(context);
    CGImageRelease(cgImage);

    out = image;
    return true;
}
#endif

}
