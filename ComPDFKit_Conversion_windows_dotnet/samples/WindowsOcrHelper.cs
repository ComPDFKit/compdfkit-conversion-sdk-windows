using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Windows.Graphics.Imaging;
using Windows.Media.Ocr;
using Windows.Storage;

namespace ComPDF_Conversion_Demo
{
    public class WindowsOcrHelper
    {
        private const string EmptyOcrResult = "{\"text_spans\":[]}";
        private string lastOcrResult = "";
        private OcrEngine ocrEngine;
        private IntPtr lastOcrResultUtf8Ptr = IntPtr.Zero;

        public WindowsOcrHelper()
        {
        }

        ~WindowsOcrHelper()
        {
            FreeUtf8Ptr();
        }

        private void FreeUtf8Ptr()
        {
            if (lastOcrResultUtf8Ptr != IntPtr.Zero)
            {
                Marshal.FreeHGlobal(lastOcrResultUtf8Ptr);
                lastOcrResultUtf8Ptr = IntPtr.Zero;
            }
        }

        public bool SetLanguage(string bcp47Tag)
        {
            try
            {
                var language = new Windows.Globalization.Language(bcp47Tag);
                if (!OcrEngine.IsLanguageSupported(language))
                {
                    var availableTags = string.Join(", ", GetAvailableLanguages());
                    System.Diagnostics.Debug.WriteLine(
                        $"[WindowsOcrHelper] OCR language '{bcp47Tag}' is not installed. " +
                        $"Available languages: [{availableTags}]. " +
                        $"Please install the language pack via: Settings > Time & Language > Language > Add a language, " +
                        $"or run 'Install-Language {bcp47Tag}' in an elevated PowerShell.");
                }

                ocrEngine = OcrEngine.TryCreateFromLanguage(language);
                if (ocrEngine == null)
                {
                    System.Diagnostics.Debug.WriteLine(
                        $"[WindowsOcrHelper] Failed to create OcrEngine for '{bcp47Tag}', falling back to user profile languages.");
                    ocrEngine = OcrEngine.TryCreateFromUserProfileLanguages();
                }

                if (ocrEngine == null)
                {
                    lastOcrResult = EmptyOcrResult;
                    return false;
                }

                return true;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine(
                    $"[WindowsOcrHelper] Windows OCR is unavailable for '{bcp47Tag}': {ex.Message}");
                ocrEngine = null;
                lastOcrResult = EmptyOcrResult;
                return false;
            }
        }

        public static List<string> GetAvailableLanguages()
        {
            try
            {
                return OcrEngine.AvailableRecognizerLanguages.Select(l => l.LanguageTag).ToList();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine(
                    $"[WindowsOcrHelper] Failed to enumerate Windows OCR languages: {ex.Message}");
                return new List<string>();
            }
        }

        public bool RunOcr(string imagePath)
        {
            try
            {
                lastOcrResult = Task.Run(async () => await RunOcrAsync(imagePath)).GetAwaiter().GetResult();
                return true;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine(
                    $"[WindowsOcrHelper] Windows OCR recognition failed: {ex.Message}");
                lastOcrResult = EmptyOcrResult;
                return false;
            }
        }

        public string GetOcrResult()
        {
            return lastOcrResult;
        }

        /// <summary>
        /// Returns a pointer to a null-terminated UTF-8 byte array in unmanaged memory.
        /// The pointer remains valid until the next call to this method or until this object is finalized.
        /// </summary>
        public IntPtr GetOcrResultUtf8Ptr()
        {
            FreeUtf8Ptr();
            byte[] utf8Bytes = Encoding.UTF8.GetBytes(lastOcrResult);
            lastOcrResultUtf8Ptr = Marshal.AllocHGlobal(utf8Bytes.Length + 1);
            Marshal.Copy(utf8Bytes, 0, lastOcrResultUtf8Ptr, utf8Bytes.Length);
            Marshal.WriteByte(lastOcrResultUtf8Ptr, utf8Bytes.Length, 0); // null terminator
            return lastOcrResultUtf8Ptr;
        }

        private async Task<string> RunOcrAsync(string imagePath)
        {
            if (ocrEngine == null)
            {
                ocrEngine = OcrEngine.TryCreateFromUserProfileLanguages();
            }
            if (ocrEngine == null)
            {
                return EmptyOcrResult;
            }

            StorageFile file = await StorageFile.GetFileFromPathAsync(imagePath);
            using (var stream = await file.OpenAsync(FileAccessMode.Read))
            {
                BitmapDecoder decoder = await BitmapDecoder.CreateAsync(stream);
                SoftwareBitmap bitmap = await decoder.GetSoftwareBitmapAsync(BitmapPixelFormat.Bgra8, BitmapAlphaMode.Premultiplied);

                OcrResult result = await ocrEngine.RecognizeAsync(bitmap);

                double textAngle = result.TextAngle.HasValue ? result.TextAngle.Value : 0.0;

                StringBuilder sb = new StringBuilder();
                sb.Append("{\"text_spans\":[");

                bool firstLine = true;
                foreach (OcrLine line in result.Lines)
                {
                    if (line.Words.Count == 0)
                        continue;

                    double lineLeft = double.MaxValue;
                    double lineTop = double.MaxValue;
                    double lineRight = double.MinValue;
                    double lineBottom = double.MinValue;

                    foreach (OcrWord word in line.Words)
                    {
                        var r = word.BoundingRect;
                        if (r.X < lineLeft) lineLeft = r.X;
                        if (r.Y < lineTop) lineTop = r.Y;
                        if (r.X + r.Width > lineRight) lineRight = r.X + r.Width;
                        if (r.Y + r.Height > lineBottom) lineBottom = r.Y + r.Height;
                    }

                    double lineHeight = lineBottom - lineTop;

                    if (!firstLine)
                        sb.Append(",");
                    firstLine = false;

                    sb.Append("{");
                    sb.AppendFormat("\"text\":{0}", EscapeJsonString(line.Text));
                    sb.AppendFormat(",\"rotation\":{0}", FormatDouble(textAngle));
                    sb.AppendFormat(",\"rect\":{{\"left\":{0},\"top\":{1},\"right\":{2},\"bottom\":{3}}}",
                        FormatDouble(lineLeft), FormatDouble(lineTop), FormatDouble(lineRight), FormatDouble(lineBottom));
                    sb.AppendFormat(",\"style\":{{\"font_size\":{0},\"font_color\":{{\"r\":0,\"g\":0,\"b\":0}}}}",
                        FormatDouble(lineHeight));

                    sb.Append(",\"words\":[");
                    bool firstWord = true;
                    foreach (OcrWord word in line.Words)
                    {
                        if (!firstWord)
                            sb.Append(",");
                        firstWord = false;

                        var wr = word.BoundingRect;
                        sb.Append("{");
                        sb.AppendFormat("\"text\":{0}", EscapeJsonString(word.Text));
                        sb.AppendFormat(",\"rect\":{{\"left\":{0},\"top\":{1},\"right\":{2},\"bottom\":{3}}}",
                            FormatDouble(wr.X), FormatDouble(wr.Y),
                            FormatDouble(wr.X + wr.Width), FormatDouble(wr.Y + wr.Height));
                        sb.Append("}");
                    }
                    sb.Append("]");

                    sb.Append("}");
                }

                sb.Append("]}");
                return sb.ToString();
            }
        }

        private static string FormatDouble(double value)
        {
            return value.ToString(System.Globalization.CultureInfo.InvariantCulture);
        }

        private static string EscapeJsonString(string s)
        {
            if (s == null) return "\"\"";
            StringBuilder sb = new StringBuilder(s.Length + 2);
            sb.Append('"');
            foreach (char c in s)
            {
                switch (c)
                {
                    case '"': sb.Append("\\\""); break;
                    case '\\': sb.Append("\\\\"); break;
                    case '\n': sb.Append("\\n"); break;
                    case '\r': sb.Append("\\r"); break;
                    case '\t': sb.Append("\\t"); break;
                    default:
                        if (c < 0x20)
                            sb.AppendFormat("\\u{0:X4}", (int)c);
                        else
                            sb.Append(c);
                        break;
                }
            }
            sb.Append('"');
            return sb.ToString();
        }

        public static string MapLanguageTag(ComPDF_Conversion.Common.OCRLanguage lang)
        {
            switch (lang)
            {
                case ComPDF_Conversion.Common.OCRLanguage.e_CHINESE: return "zh-Hans";
                case ComPDF_Conversion.Common.OCRLanguage.e_CHINESE_TRA: return "zh-Hant";
                case ComPDF_Conversion.Common.OCRLanguage.e_ENGLISH: return "en-US";
                case ComPDF_Conversion.Common.OCRLanguage.e_KOREAN: return "ko";
                case ComPDF_Conversion.Common.OCRLanguage.e_JAPANESE: return "ja";
                case ComPDF_Conversion.Common.OCRLanguage.e_LATIN: return "la";
                case ComPDF_Conversion.Common.OCRLanguage.e_DEVANAGARI: return "hi";
                case ComPDF_Conversion.Common.OCRLanguage.e_Cyrillic: return "ru";
                case ComPDF_Conversion.Common.OCRLanguage.e_Arabic: return "ar";
                case ComPDF_Conversion.Common.OCRLanguage.e_Tamil: return "ta";
                case ComPDF_Conversion.Common.OCRLanguage.e_Telugu: return "te";
                case ComPDF_Conversion.Common.OCRLanguage.e_Kannada: return "kn";
                case ComPDF_Conversion.Common.OCRLanguage.e_Thai: return "th";
                case ComPDF_Conversion.Common.OCRLanguage.e_Greek: return "el";
                case ComPDF_Conversion.Common.OCRLanguage.e_Eslav: return "sr";
                default: return "en-US";
            }
        }
    }
}
