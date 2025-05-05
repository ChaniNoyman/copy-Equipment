using System;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Linq;
using System.Collections.Generic;

public class CsvProcessor
{
    public static string destinationLogFile = @"C:\Users\Chanif\Desktop\C#\CopyExcelFileEquipment\Log.csv";
    private static readonly string[] searchPhrases = { "ציוד נדרש", "ציוד דרוש", "ציוד:", "ציוד נדרש למדריכים" };
    private static readonly Regex searchRegex = new Regex(string.Join("|", searchPhrases.Select(Regex.Escape)));
    private static readonly Regex lineSplitRegex = new Regex(@"(?<!\r)\n"); // פיצול רק לפי LF שלא מגיע אחרי CR
    private static readonly Regex htmlTagRegex = new Regex(@"<(?!\/?strong\b)([^>]+)>"); // תגית HTML פותחת או סוגרת שאינה STRONG

    // פונקציה להסרת תגיות HTML ממחרוזת
    private static string RemoveHtmlTags(string input)
    {
        string withoutNbsp = input.Replace("&nbsp;", " ");
        return Regex.Replace(withoutNbsp, "<.*?>", string.Empty);
    }

    private static void LogNoCommas(string columnA, string problematicText)
    {
        try
        {
            Console.WriteLine(problematicText);
            string cleanedText = RemoveHtmlTags(problematicText);
            File.AppendAllText(destinationLogFile, $"{columnA}\t{cleanedText}\n"); // סוף שורה LF
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
            if (textToSplit[i] == ',' &&
                !textToSplit.Substring(0, i).Contains('.') &&
                !textToSplit.Substring(0, i).Contains("מהלך הפעולה") &&
                !textToSplit.Substring(0, i).Contains("!#!#"))
            {
                string part = textToSplit.Substring(startIndex, i - startIndex).Trim();
                if (part.StartsWith(":") || part.StartsWith("-"))
                {
                    part = part.Substring(1).Trim();
                }
                parts.Add(RemoveHtmlTags(part));
                startIndex = i + 1;
            }
        }
        string lastPart = textToSplit.Substring(startIndex).Trim();
        if (lastPart.StartsWith(":") || lastPart.StartsWith("-"))
        {
            lastPart = lastPart.Substring(1).Trim();
        }
        parts.Add(RemoveHtmlTags(lastPart));
        return string.Join("\t", parts); // עדיין משתמש בטאב כהפרדה פנימית לאחר הפיצול
    }

    public static void ProcessCsvReadAllText(string sourceFilePath, string destinationFilePath)
    {
        try
        {
            string allText = File.ReadAllText(sourceFilePath, Encoding.UTF8);
            string[] lines = lineSplitRegex.Split(allText); // פיצול רק לפי LF שלא מגיע אחרי CR

            using (var writer = new StreamWriter(destinationFilePath, false, Encoding.UTF8))
            {
                writer.NewLine = "\n"; // הגדרת סוף שורה ל-LF באופן מפורש
                foreach (string line in lines)
                {
                    string cleanedLine = line.Trim().Replace("\r", ""); // הסרת CR אם נשאר בסוף שורה
                    string[] columns = cleanedLine.Split(',');

                    if (columns.Length >= 1)
                    {
                        string columnA = columns[0];
                        string remainingContent = columns.Length > 1 ? string.Join(",", columns.Skip(1)) : "";
                        string cleanedContent = RemoveHtmlTags(remainingContent);
                        string extractedEquipmentString = "";

                        MatchCollection matches = searchRegex.Matches(cleanedContent);

                        if (matches.Count > 0)
                        {
                            extractedEquipmentString = string.Join("\t", matches.Select(match =>
                            {
                                string textAfterPhrase = cleanedContent.Substring(match.Index + match.Length).Trim();

                                // מציאת המיקום של תגית HTML פותחת/סוגרת חדשה שאינה STRONG
                                Match htmlMatch = htmlTagRegex.Match(textAfterPhrase);
                                int indexHtmlTag = htmlMatch.Success ? htmlMatch.Index : -1;

                                int indexDot = textAfterPhrase.IndexOf('.');
                                int indexMahaloch = textAfterPhrase.IndexOf("מהלך הפעולה");
                                int indexHashtags = textAfterPhrase.IndexOf("!#!#");

                                int endIndex = textAfterPhrase.Length;
                                if (indexDot != -1 && indexDot < endIndex) endIndex = indexDot;
                                if (indexMahaloch != -1 && indexMahaloch < endIndex) endIndex = indexMahaloch;
                                if (indexHashtags != -1 && indexHashtags < endIndex) endIndex = indexHashtags;
                                if (indexHtmlTag != -1 && indexHtmlTag < endIndex) endIndex = indexHtmlTag; // הוספת התנאי החדש

                                string textToSplit = textAfterPhrase.Substring(0, endIndex).Trim();
                                string cleanedAndSplit = SplitAndClean(textToSplit);
                                return string.Join("\t", cleanedAndSplit.Split('\t'));
                            }).ToArray());
                        }

                        writer.WriteLine(columnA + (string.IsNullOrEmpty(extractedEquipmentString) ? "" : "\t" + extractedEquipmentString));
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

        ProcessCsvReadAllText(sourceFile, destinationFile); // ודא שאתה משתמש במתודה הזו
        Console.ReadKey();
    }
}