using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace fix_jpdotx_for_pandoc
{
    internal class Program
    {
        static void Main(string[] args)
        {
            string sourceFilePath = @"docx-style.dotx";
            string destinationFilePath = @"docx-style.dotx";

            // 置換前後のスタイル名を辞書として定義
            var styleNameMappings = new Dictionary<string, string>
            {
                { "標準", "Normal" },
                { "本文", "Body Text" },
                { "本文 (文字)", "Body Text Char" },
                // First Paragraph // ローカライズされないので無処理で問題なし
                // Compact // ローカライズされないので無処理で問題なし
                { "表題", "Title" },
                { "表題 (文字)", "Title Char" },
                { "副題", "Subtitle" },
                { "副題 (文字)", "Subtitle Char" },
                // Author // ローカライズされないので無処理で問題なし
                { "日付", "Date" },
                { "日付 (文字)", "Date Char" },
                // Abstract Title // ローカライズされないので無処理で問題なし
                // Abstract // ローカライズされないので無処理で問題なし
                { "文献目録", "Bibliography" },
                { "見出し 1", "heading 1" },
                { "見出し 1 (文字)", "Heading 1 Char" },
                { "見出し 2", "heading 2" },
                { "見出し 2 (文字)", "Heading 2 Char" },
                { "見出し 3", "heading 3" },
                { "見出し 3 (文字)", "Heading 3 Char" },
                { "見出し 4", "heading 4" },
                { "見出し 4 (文字)", "Heading 4 Char" },
                { "見出し 5", "heading 5" },
                { "見出し 5 (文字)", "Heading 5 Char" },
                { "見出し 6", "heading 6" },
                { "見出し 6 (文字)", "Heading 6 Char" },
                { "見出し 7", "heading 7" },
                { "見出し 7 (文字)", "Heading 7 Char" },
                { "見出し 8", "heading 8" },
                { "見出し 8 (文字)", "Heading 8 Char" },
                { "見出し 9", "heading 9" },
                { "見出し 9 (文字)", "Heading 9 Char" },
                { "ブロック", "Block Text" },
                { "脚注文字列", "Footnote Text" },
                // Footnote Block Text // ローカライズされないので無処理で問題なし
                { "段落フォント", "Default Paragraph Font" },
                // Table // ローカライズされないので無処理で問題なし
                // Definition Term // ローカライズされないので無処理で問題なし
                // Definition // ローカライズされないので無処理で問題なし
                { "図表番号", "Caption" },
                // Table Caption // ローカライズされないので無処理で問題なし
                // Image Caption // ローカライズされないので無処理で問題なし
                // Figure // ローカライズされないので無処理で問題なし
                // Captioned Figure // ローカライズされないので無処理で問題なし
                // Verbatim Char // ローカライズされないので無処理で問題なし
                // Section Number // ローカライズされないので無処理で問題なし
                { "脚注参照", "Footnote Reference" },
                { "ハイパーリンク", "Hyperlink" },
                { "目次の見出し", "TOC Heading" },
                { "ヘッダー (文字)", "header Char" }, // 必須ではない
                { "フッター (文字)", "footer Char" }, // 必須ではない
            };

            try
            {
                // 元のファイルをコピーして新しいファイルに保存
                if (sourceFilePath != destinationFilePath)
                {
                    File.Copy(sourceFilePath, destinationFilePath, overwrite: true);
                }

                // 新しいファイルを読み書きモードで開く
                using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(destinationFilePath, true))
                {
                    var styleDefinitionsPart = wordDoc.MainDocumentPart?.StyleDefinitionsPart;

                    if (styleDefinitionsPart == null || styleDefinitionsPart.Styles == null)
                    {
                        Console.WriteLine("スタイル定義が見つかりません。");
                        return;
                    }

                    var styles = styleDefinitionsPart.Styles.Elements<Style>().ToList();

                    Console.WriteLine("=== LIST STYLES (BEFORE) START ===");
                    Console.WriteLine("Style\tID\tBasedOn\tNextParagraphStyle\tLinkedStyle");
                    foreach (var style in styles)
                    {
                        Console.WriteLine($"{style.StyleName?.Val}\t{style.StyleId}\t{style.BasedOn?.Val}\t{style.NextParagraphStyle?.Val}\t{style.LinkedStyle?.Val}");
                    }
                    Console.WriteLine("=== LIST STYLES (BEFORE) END ===");

                    Console.WriteLine("=== PATCH START ===");
                    foreach (var mapping in styleNameMappings)
                    {
                        string oldStyleName = mapping.Key;
                        string newStyleName = mapping.Value;
                        string newStyleID = char.ToUpper(newStyleName.Replace(" ", "")[0]) + newStyleName.Replace(" ", "").Substring(1);

                        // 置換対象のスタイルを検索
                        var targetStyle = styles.FirstOrDefault(style => style.StyleName?.Val == oldStyleName);
                        if (targetStyle == null)
                        {
                            targetStyle = styles.FirstOrDefault(style => style.StyleName?.Val == newStyleName);
                        }

                        if (targetStyle != null)
                        {
                            //Console.WriteLine($"スタイル '{targetStyle.StyleName?.Val}' を見つけました。");

                            if (targetStyle.StyleName?.Val != newStyleName)
                            {
                                // スタイル名を置換 (実はいらないかも)
                                Console.WriteLine($"スタイル名を '{targetStyle.StyleName?.Val}' から '{newStyleName}' に変更します。");

                                // 競合チェック: newStyleName が既存のスタイル名と競合する場合スキップ
                                if (styles.Any(style => style.StyleName?.Val == newStyleName))
                                {
                                    Console.WriteLine($"競合エラー: 新しいスタイル名 '{newStyleName}' は既存のスタイル名と競合しています。スキップします。");
                                    continue;
                                }

#pragma warning disable CS8602 // null 参照の可能性があるものの逆参照です。
                                targetStyle.StyleName.Val = newStyleName;
#pragma warning restore CS8602 // null 参照の可能性があるものの逆参照です。
                            }

                            // すでにスタイルIDが期待通りなら何もしない
                            if (targetStyle.StyleId == newStyleID)
                            {
                                continue;
                            }

                            // スタイルIDを置換
                            Console.WriteLine($"スタイルIDを '{targetStyle.StyleId}' から '{newStyleID}' に変更します。");

                            // 競合チェック: newStyleID が既存のスタイル名と競合する場合スキップ
                            if (styles.Any(style => style.StyleId == newStyleID))
                            {
                                Console.WriteLine($"競合エラー: 新しいスタイルID '{newStyleID}' は既存のスタイルIDと競合しています。スキップします。");
                                continue;
                            }

                            var oldStyleId = targetStyle.StyleId;
                            targetStyle.StyleId = newStyleID;

                            // BasedOn を更新
                            foreach (var style in styles)
                            {
                                if (style.BasedOn?.Val == oldStyleId)
                                {
                                    Console.WriteLine($"継承関係を更新: スタイル '{style.StyleName?.Val}' の BasedOn を '{style.BasedOn?.Val}' から '{targetStyle.StyleId}' に更新します。");
#pragma warning disable CS8602 // null 参照の可能性があるものの逆参照です。
                                    style.BasedOn.Val = targetStyle.StyleId;
#pragma warning restore CS8602 // null 参照の可能性があるものの逆参照です。
                                }
                            }

                            // NextParagraphStyle を更新
                            foreach (var style in styles)
                            {
                                if (style.NextParagraphStyle?.Val == oldStyleId)
                                {
                                    Console.WriteLine($"継承関係を更新: スタイル '{style.StyleName?.Val}' の NextParagraphStyle を '{style.NextParagraphStyle?.Val}' から '{targetStyle.StyleId}' に更新します。");
#pragma warning disable CS8602 // null 参照の可能性があるものの逆参照です。
                                    style.NextParagraphStyle.Val = targetStyle.StyleId;
#pragma warning restore CS8602 // null 参照の可能性があるものの逆参照です。
                                }
                            }

                            // LinkedStyle を更新
                            foreach (var style in styles)
                            {
                                if (style.LinkedStyle?.Val == oldStyleId)
                                {
                                    Console.WriteLine($"継承関係を更新: スタイル '{style.StyleName?.Val}' の LinkedStyle を '{style.LinkedStyle?.Val}' から '{targetStyle.StyleId}' に更新します。");
#pragma warning disable CS8602 // null 参照の可能性があるものの逆参照です。
                                    style.LinkedStyle.Val = targetStyle.StyleId;
#pragma warning restore CS8602 // null 参照の可能性があるものの逆参照です。
                                }
                            }
                        }
                    }
                    Console.WriteLine("=== PATCH END ===");

                    Console.WriteLine("=== LIST STYLES (AFTER) START ===");
                    Console.WriteLine("Style\tID\tBasedOn\tNextParagraphStyle\tLinkedStyle");
                    foreach (var style in styles)
                    {
                        Console.WriteLine($"{style.StyleName?.Val}\t{style.StyleId}\t{style.BasedOn?.Val}\t{style.NextParagraphStyle?.Val}\t{style.LinkedStyle?.Val}");
                    }
                    Console.WriteLine("=== LIST STYLES (AFTER) END ===");

                    // 保存
                    styleDefinitionsPart.Styles.Save();
                    Console.WriteLine($"スタイル名の変更が完了しました。変更内容は '{destinationFilePath}' に保存されました。");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"エラー: {ex.Message}");
            }
        }
    }
}
