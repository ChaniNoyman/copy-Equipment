using System;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Linq;
using System.Collections.Generic;

public class CsvProcessor
{
    public static string destinationLogFile = @"C:\Users\Chanif\Desktop\C#\CopyExcelFileEquipment\Log.csv";
    private static readonly string[] searchPhrases = { "ציוד נדרש", "ציוד דרוש", "ציוד:", "ציוד נדרש למדריכים", "ציוד נדרש לפעולה", "הכנות וציוד נדרש" };
    private static readonly Regex searchRegex = new Regex(string.Join("|", searchPhrases.Select(Regex.Escape)));
    // regex זה לא יהיה בשימוש ישיר לניקוי, כי נשתמש ב-RemoveHtmlTags בצורה גורפת יותר
    // אבל הוא עדיין שימושי אם תרצה לבדוק תגיות מסוימות בנפרד (כמו ב-SplitAndClean).
    private static readonly Regex htmlTagRegex = new Regex(@"<(?!\/?(?:strong|b)\b)([^>]+)>");

    private static readonly string[] stopPhrases = { ".", "מהלך הפעולה", "!#!#", @"""," };

    // פונקציה להסרת תגיות HTML ממחרוזת
    private static string RemoveHtmlTags(string input)
    {
        string withoutNbsp = input.Replace("&nbsp;", " ");
        return Regex.Replace(withoutNbsp, "<.*?>", string.Empty); // מסיר את כל תגיות ה-HTML
    }

    private static void LogNoCommas(string columnA, string problematicText)
    {
        try
        {
            Console.WriteLine(problematicText);
            string cleanedText = RemoveHtmlTags(problematicText);
            File.AppendAllText(destinationLogFile, $"{columnA}\t{cleanedText}\n");
            Console.WriteLine("after: " + cleanedText);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"שגיאה בכתיבה ללוג: {ex.Message}");
        }
    }

    private static string SplitAndClean(string textToSplit)
    {
        List<string> parts = new List<string>();
        int startIndex = 0;
        for (int i = 0; i < textToSplit.Length; i++)
        {
            // התנאים האלה עדיין טובים, גם אם הטקסט כבר נקי מ-HTML,
            // הם פשוט לא ימצאו תגיות.
            if (textToSplit[i] == ',' &&
                !textToSplit.Substring(0, i).Contains('.') &&
                !textToSplit.Substring(0, i).Contains("מהלך הפעולה") &&
                !textToSplit.Substring(0, i).Contains("!#!#")) // אין צורך ב-htmlTagRegex.IsMatch כאן, כי כבר ניקינו HTML
            {
                string part = textToSplit.Substring(startIndex, i - startIndex).Trim();
                if (part.StartsWith(":") || part.StartsWith("-"))
                {
                    part = part.Substring(1).Trim();
                }
                parts.Add(part); // אין צורך ב-RemoveHtmlTags כאן שוב, כי הטקסט כבר נקי
                startIndex = i + 1;
            }
        }
        string lastPart = textToSplit.Substring(startIndex).Trim();
        if (lastPart.StartsWith(":") || lastPart.StartsWith("-"))
        {
            lastPart = lastPart.Substring(1).Trim();
        }
        parts.Add(lastPart); // אין צורך ב-RemoveHtmlTags כאן שוב
        return string.Join("\t", parts);
    }

    public static void ProcessCsvReadAllText(string sourceFilePath, string destinationFilePath)
    {
        try
        {
            string allText = File.ReadAllText(sourceFilePath, Encoding.UTF8);
            string[] lines = new Regex(@"(?<!\r)\n").Split(allText);

            using (var writer = new StreamWriter(destinationFilePath, false, Encoding.UTF8))
            {
                writer.NewLine = "\n";
                foreach (string line in lines)
                {
                    string cleanedLine = line.Trim().Replace("\r", "");

                    if (string.IsNullOrWhiteSpace(cleanedLine.Replace("\t", "").Replace("\u00A0", "")))
                    {
                        continue;
                    }

                    string[] columns = cleanedLine.Split(',');

                    if (columns.Length >= 1)
                    {
                        string columnA = columns[0];
                        string originalRemainingContent = columns.Length > 1 ? string.Join(",", columns.Skip(1)) : "";

                        // **השינוי המרכזי כאן:** נקה את כל ה-HTML מ-remainingContent בתחילת התהליך
                        string cleanRemainingContent = RemoveHtmlTags(originalRemainingContent);

                        string extractedEquipmentString = "";

                        // כעת, החיפוש יתבצע על הטקסט הנקי מ-HTML
                        MatchCollection matches = searchRegex.Matches(cleanRemainingContent);

                        if (matches.Count > 0)
                        {
                            extractedEquipmentString = string.Join("\t", matches.Select(match =>
                            {
                                // התחלה של הטקסט לאחר הביטוי המבוקש ("ציוד נדרש:") בטקסט הנקי.
                                // מכיוון שהטקסט כבר נקי מ-HTML, אין צורך לדלג על תגיות <strong>/<b> ספציפיות.
                                string currentSegment = cleanRemainingContent.Substring(match.Index + match.Length).Trim();

                                int endIndex = currentSegment.Length;

                                // חיפוש המיקום הראשון של אחת ממילות העצירה
                                foreach (string stopPhrase in stopPhrases)
                                {
                                    int index = currentSegment.IndexOf(stopPhrase);
                                    if (index != -1 && index < endIndex)
                                    {
                                        endIndex = index;
                                    }
                                }

                                // מכיוון ש-currentSegment כבר נקי מ-HTML, בדיקה זו של htmlMatch מיותרת.
                                // אך אם תרצה להשאיר אותה למקרה של תגיות HTML שעדיין נותרו מסיבה כלשהי, זה לא יזיק.
                                // כרגע, היא פשוט לא תמצא כלום אם RemoveHtmlTags עבד כראוי.
                                // Match htmlMatch = htmlTagRegex.Match(currentSegment);
                                // if (htmlMatch.Success && htmlMatch.Index < endIndex)
                                // {
                                //     endIndex = htmlMatch.Index;
                                // }

                                string textToSplit = currentSegment.Substring(0, endIndex).Trim();
                                string cleanedAndSplit = SplitAndClean(textToSplit);
                                return string.Join("\t", cleanedAndSplit.Split('\t'));
                            }).ToArray());
                        }

                        writer.WriteLine(columnA + (string.IsNullOrEmpty(extractedEquipmentString) ? "" : "\t" + extractedEquipmentString.Replace("\n", " ")));
                    }
                }
            }
            Console.WriteLine("העיבוד הסתיים...");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"שגיאה: {ex.Message}");
        }
    }

    public static void Main(string[] args)
    {
        string sourceFile = @"C:\Users\Chanif\Desktop\C#\CopyExcelFileEquipment\data.csv";
        string destinationFile = @"C:\Users\Chanif\Desktop\C#\CopyExcelFileEquipment\fixDataE.csv";

        ProcessCsvReadAllText(sourceFile, destinationFile);
        Console.ReadKey();
    }
}