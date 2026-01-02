using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Wxy.Xaf.ExcelExchange.Services
{
    /// <summary>
    /// 文件编码检测工具类
    /// </summary>
    public static class EncodingDetector
    {
        /// <summary>
        /// 检测文件内容的编码格式
        /// </summary>
        /// <param name="content">文件内容字节数组</param>
        /// <param name="maxBytesToCheck">最大检测字节数（默认8KB）</param>
        /// <returns>编码检测结果</returns>
        public static EncodingDetectionResult DetectEncoding(byte[] content, int maxBytesToCheck = 8192)
        {
            var result = new EncodingDetectionResult();
            
            if (content == null || content.Length == 0)
            {
                result.DetectedEncoding = Encoding.UTF8;
                result.Confidence = 0;
                result.DetectionMethod = "Empty file - default to UTF-8";
                return result;
            }

            // 限制检测范围以提高性能
            int bytesToCheck = Math.Min(content.Length, maxBytesToCheck);
            byte[] sampleContent = content.Take(bytesToCheck).ToArray();

            // Phase 1: BOM检测（最高优先级）
            var bomResult = DetectBom(content);
            if (bomResult.HasBom)
            {
                result.DetectedEncoding = bomResult.Encoding;
                result.Confidence = 100;
                result.DetectionMethod = $"BOM detected: {bomResult.BomType}";
                result.HasBom = true;
                result.TriedEncodings.Add($"BOM-{bomResult.BomType}");
                return result;
            }

            // Phase 2: 强制尝试所有可能的编码，寻找最佳中文字符匹配
            var encodingCandidates = GetEncodingCandidates();
            double bestChineseRatio = 0;
            Encoding bestEncoding = null;
            string bestMethod = null;

            foreach (var candidate in encodingCandidates)
            {
                try
                {
                    string decodedText = candidate.Encoding.GetString(sampleContent);
                    double chineseRatio = CalculateChineseCharacterRatio(decodedText);
                    
                    result.TriedEncodings.Add($"{candidate.Name} (Chinese ratio: {chineseRatio:F1}%)");

                    if (chineseRatio > bestChineseRatio)
                    {
                        bestChineseRatio = chineseRatio;
                        bestEncoding = candidate.Encoding;
                        bestMethod = $"Chinese character analysis: {candidate.Name}";
                    }
                }
                catch (Exception ex)
                {
                    result.TriedEncodings.Add($"{candidate.Name} (failed: {ex.Message})");
                }
            }

            // 如果找到了包含中文字符的编码，使用它
            if (bestChineseRatio >= 5) // 降低阈值到5%
            {
                result.DetectedEncoding = bestEncoding;
                result.Confidence = bestChineseRatio;
                result.DetectionMethod = bestMethod;
                return result;
            }

            // Phase 3: UTF-8验证（如果没有找到中文字符）
            var utf8Result = ValidateUtf8(sampleContent);
            result.TriedEncodings.Add($"UTF-8-NoBOM (score: {utf8Result.Score:F1}%)");
            
            if (utf8Result.IsValid && utf8Result.Score >= 80)
            {
                result.DetectedEncoding = Encoding.UTF8;
                result.Confidence = utf8Result.Score;
                result.DetectionMethod = "UTF-8 byte sequence validation";
                return result;
            }

            // Phase 4: 回退策略
            if (utf8Result.Score >= 50)
            {
                result.DetectedEncoding = Encoding.UTF8;
                result.Confidence = utf8Result.Score;
                result.DetectionMethod = "UTF-8 fallback (partial validation)";
                return result;
            }

            // 最终回退到系统默认编码
            result.DetectedEncoding = Encoding.Default;
            result.Confidence = 30;
            result.DetectionMethod = "System default encoding fallback";
            result.TriedEncodings.Add("System-Default-Fallback");

            return result;
        }

        /// <summary>
        /// 获取编码候选列表
        /// </summary>
        private static List<EncodingCandidate> GetEncodingCandidates()
        {
            var candidates = new List<EncodingCandidate>();

            // 添加UTF-8
            candidates.Add(new EncodingCandidate { Name = "UTF-8", Encoding = Encoding.UTF8 });

            // 尝试添加中文编码
            var chineseEncodings = new[]
            {
                ("GB2312", "GB2312"),
                ("GBK", "GBK"), 
                ("GB18030", "GB18030"),
                ("Big5", "Big5"),
                ("Windows-936", "936"), // 简体中文
                ("Windows-950", "950")  // 繁体中文
            };

            foreach (var (name, codePage) in chineseEncodings)
            {
                try
                {
                    var encoding = Encoding.GetEncoding(codePage);
                    candidates.Add(new EncodingCandidate { Name = name, Encoding = encoding });
                }
                catch
                {
                    // 编码不可用，跳过
                }
            }

            // 添加其他常用编码
            try
            {
                candidates.Add(new EncodingCandidate { Name = "Windows-1252", Encoding = Encoding.GetEncoding("Windows-1252") });
            }
            catch
            {
                // Windows-1252 编码不可用,忽略
                System.Diagnostics.Debug.WriteLine("[EncodingDetector] Windows-1252 encoding not available");
            }

            // 添加系统默认编码
            if (Encoding.Default != Encoding.UTF8)
            {
                candidates.Add(new EncodingCandidate { Name = "System-Default", Encoding = Encoding.Default });
            }

            return candidates;
        }

        /// <summary>
        /// 编码候选项
        /// </summary>
        private class EncodingCandidate
        {
            public string Name { get; set; }
            public Encoding Encoding { get; set; }
        }

        /// <summary>
        /// 检测字节顺序标记(BOM)
        /// </summary>
        private static BomDetectionResult DetectBom(byte[] content)
        {
            if (content.Length >= 3)
            {
                // UTF-8 BOM: EF BB BF
                if (content[0] == 0xEF && content[1] == 0xBB && content[2] == 0xBF)
                {
                    return new BomDetectionResult
                    {
                        HasBom = true,
                        Encoding = Encoding.UTF8,
                        BomType = "UTF-8",
                        BomLength = 3
                    };
                }
            }

            if (content.Length >= 2)
            {
                // UTF-16 BE BOM: FE FF
                if (content[0] == 0xFE && content[1] == 0xFF)
                {
                    return new BomDetectionResult
                    {
                        HasBom = true,
                        Encoding = Encoding.BigEndianUnicode,
                        BomType = "UTF-16-BE",
                        BomLength = 2
                    };
                }

                // UTF-16 LE BOM: FF FE
                if (content[0] == 0xFF && content[1] == 0xFE)
                {
                    // 需要检查是否是UTF-32 LE BOM
                    if (content.Length >= 4 && content[2] == 0x00 && content[3] == 0x00)
                    {
                        return new BomDetectionResult
                        {
                            HasBom = true,
                            Encoding = Encoding.UTF32,
                            BomType = "UTF-32-LE",
                            BomLength = 4
                        };
                    }

                    return new BomDetectionResult
                    {
                        HasBom = true,
                        Encoding = Encoding.Unicode,
                        BomType = "UTF-16-LE",
                        BomLength = 2
                    };
                }
            }

            if (content.Length >= 4)
            {
                // UTF-32 BE BOM: 00 00 FE FF
                if (content[0] == 0x00 && content[1] == 0x00 && content[2] == 0xFE && content[3] == 0xFF)
                {
                    return new BomDetectionResult
                    {
                        HasBom = true,
                        Encoding = Encoding.GetEncoding("UTF-32BE"),
                        BomType = "UTF-32-BE",
                        BomLength = 4
                    };
                }
            }

            return new BomDetectionResult { HasBom = false };
        }

        /// <summary>
        /// 验证UTF-8字节序列
        /// </summary>
        private static Utf8ValidationResult ValidateUtf8(byte[] content)
        {
            int validBytes = 0;
            int totalBytes = content.Length;
            int invalidSequences = 0;

            for (int i = 0; i < content.Length; i++)
            {
                byte b = content[i];

                // ASCII字符 (0x00-0x7F)
                if (b <= 0x7F)
                {
                    validBytes++;
                    continue;
                }

                // 多字节UTF-8序列
                int expectedBytes = 0;
                if ((b & 0xE0) == 0xC0) expectedBytes = 2;      // 110xxxxx
                else if ((b & 0xF0) == 0xE0) expectedBytes = 3; // 1110xxxx
                else if ((b & 0xF8) == 0xF0) expectedBytes = 4; // 11110xxx
                else
                {
                    invalidSequences++;
                    continue;
                }

                // 检查后续字节
                bool validSequence = true;
                if (i + expectedBytes > content.Length)
                {
                    validSequence = false;
                }
                else
                {
                    for (int j = 1; j < expectedBytes; j++)
                    {
                        if ((content[i + j] & 0xC0) != 0x80) // 10xxxxxx
                        {
                            validSequence = false;
                            break;
                        }
                    }
                }

                if (validSequence)
                {
                    validBytes += expectedBytes;
                    i += expectedBytes - 1; // -1因为循环会自动+1
                }
                else
                {
                    invalidSequences++;
                }
            }

            double score = totalBytes > 0 ? (double)validBytes / totalBytes * 100 : 0;
            bool isValid = score >= 80 && invalidSequences < totalBytes * 0.1; // 调整验证条件

            return new Utf8ValidationResult
            {
                IsValid = isValid,
                Score = score,
                ValidBytes = validBytes,
                InvalidSequences = invalidSequences
            };
        }

        /// <summary>
        /// 检测中文编码
        /// </summary>
        private static ChineseEncodingResult DetectChineseEncoding(byte[] content)
        {
            var result = new ChineseEncodingResult();
            var encodingsToTry = new List<(string Name, Encoding Encoding)>();

            // 安全地获取编码，避免不支持的编码导致异常
            try
            {
                encodingsToTry.Add(("GB2312", Encoding.GetEncoding("GB2312")));
            }
            catch
            {
                result.TriedEncodings.Add("GB2312 (not available)");
            }

            try
            {
                encodingsToTry.Add(("GBK", Encoding.GetEncoding("GBK")));
            }
            catch
            {
                result.TriedEncodings.Add("GBK (not available)");
            }

            try
            {
                encodingsToTry.Add(("Big5", Encoding.GetEncoding("Big5")));
            }
            catch
            {
                result.TriedEncodings.Add("Big5 (not available)");
            }

            double bestConfidence = 0;
            Encoding bestEncoding = null;
            string bestEncodingName = null;

            foreach (var encodingInfo in encodingsToTry)
            {
                try
                {
                    string decodedText = encodingInfo.Encoding.GetString(content);
                    double chineseRatio = CalculateChineseCharacterRatio(decodedText);
                    
                    result.TriedEncodings.Add($"{encodingInfo.Name} (Chinese ratio: {chineseRatio:F1}%)");

                    if (chineseRatio > bestConfidence)
                    {
                        bestConfidence = chineseRatio;
                        bestEncoding = encodingInfo.Encoding;
                        bestEncodingName = encodingInfo.Name;
                    }
                }
                catch
                {
                    result.TriedEncodings.Add($"{encodingInfo.Name} (decode failed)");
                }
            }

            result.BestEncoding = bestEncoding;
            result.Confidence = bestConfidence;
            result.EncodingName = bestEncodingName;

            return result;
        }

        /// <summary>
        /// 计算中文字符比例
        /// </summary>
        public static double CalculateChineseCharacterRatio(string text)
        {
            if (string.IsNullOrEmpty(text))
                return 0;

            int chineseCount = 0;
            int totalCount = 0;

            foreach (char c in text)
            {
                // 跳过控制字符和空白字符
                if (char.IsControl(c) || char.IsWhiteSpace(c))
                    continue;

                totalCount++;

                if (IsChineseCharacter(c))
                {
                    chineseCount++;
                }
            }

            return totalCount > 0 ? (double)chineseCount / totalCount * 100 : 0;
        }

        /// <summary>
        /// 判断是否为中文字符
        /// </summary>
        public static bool IsChineseCharacter(char c)
        {
            int code = (int)c;
            
            // CJK统一汉字 (U+4E00-U+9FFF)
            if (code >= 0x4E00 && code <= 0x9FFF) return true;
            
            // CJK扩展A (U+3400-U+4DBF)
            if (code >= 0x3400 && code <= 0x4DBF) return true;
            
            // CJK符号和标点 (U+3000-U+303F)
            if (code >= 0x3000 && code <= 0x303F) return true;
            
            // CJK兼容 (U+3300-U+33FF)
            if (code >= 0x3300 && code <= 0x33FF) return true;
            
            // 全角ASCII、全角中英文标点 (U+FF00-U+FFEF)
            if (code >= 0xFF00 && code <= 0xFFEF) return true;
            
            return false;
        }

        /// <summary>
        /// 检查文本是否包含中文字符
        /// </summary>
        public static bool ContainsChineseCharacters(string text)
        {
            if (string.IsNullOrEmpty(text))
                return false;

            return text.Any(IsChineseCharacter);
        }

        #region 内部类

        private class BomDetectionResult
        {
            public bool HasBom { get; set; }
            public Encoding Encoding { get; set; }
            public string BomType { get; set; }
            public int BomLength { get; set; }
        }

        private class Utf8ValidationResult
        {
            public bool IsValid { get; set; }
            public double Score { get; set; }
            public int ValidBytes { get; set; }
            public int InvalidSequences { get; set; }
        }

        private class ChineseEncodingResult
        {
            public Encoding BestEncoding { get; set; }
            public double Confidence { get; set; }
            public string EncodingName { get; set; }
            public List<string> TriedEncodings { get; set; } = new List<string>();
        }

        #endregion
    }

    /// <summary>
    /// 编码检测结果
    /// </summary>
    public class EncodingDetectionResult
    {
        /// <summary>
        /// 检测到的编码
        /// </summary>
        public Encoding DetectedEncoding { get; set; }

        /// <summary>
        /// 检测置信度 (0-100)
        /// </summary>
        public double Confidence { get; set; }

        /// <summary>
        /// 检测方法描述
        /// </summary>
        public string DetectionMethod { get; set; }

        /// <summary>
        /// 是否包含BOM
        /// </summary>
        public bool HasBom { get; set; }

        /// <summary>
        /// 尝试过的编码列表（用于调试）
        /// </summary>
        public List<string> TriedEncodings { get; set; } = new List<string>();

        /// <summary>
        /// 返回检测结果的字符串表示
        /// </summary>
        public override string ToString()
        {
            return $"Encoding: {DetectedEncoding?.EncodingName}, Confidence: {Confidence:F1}%, Method: {DetectionMethod}";
        }
    }
}