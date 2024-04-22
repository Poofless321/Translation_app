using Newtonsoft.Json;
using OfficeOpenXml;
using System.IO;
using System.Security;
using TranslatorApp;


class Translator
{
    static readonly HttpClient client = new HttpClient();

    public static bool IsJapanese(string word)
    {
        foreach (char c in word)
        {
            if ((c >= '\u3040' && c <= '\u30FF') ||
                (c >= '\u3400' && c <= '\u4DBF') ||
                (c >= '\u4E00' && c <= '\u9FFF') ||
                (c >= '\uF900' && c <= '\uFAFF'))
            {
                return true;
            }
        }
        return false;
    }

    public static async Task<string?> API(string text)
    {
        //give the fucker live
        API_DATA OpenAI = new API_DATA("https://api.openai.com/v1/chat/completions", TimeSpan.FromMinutes(15));
        OpenAI.headers(); // sets headers

        return await OpenAI.API(text);
    }

    public static void ProgressLog(string?[] text)
    {
        foreach (var batch in text)
        {
            try
            {
                // Attempt to deserialize the JSON string
                var jsonObject = JsonConvert.DeserializeObject(batch);
                if (jsonObject == null)
                {
                    Console.WriteLine("Warning: Deserialization returned null.");
                    continue;
                }

                // Attempt to serialize the object back to a JSON string
                string currentBatch = JsonConvert.SerializeObject(jsonObject, Formatting.Indented);

                Console.WriteLine("=-=-=-=-=-=-=-=-=-=-=-=-=-=-==-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-");
                Console.WriteLine(currentBatch);
                Console.WriteLine("=-=-=-=-=-=-=-=-=-=-=-=-=-=-==-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-");
                Console.WriteLine();
            }
            catch (JsonReaderException jsonReaderEx)
            {
                // Handle JSON formatting issues
                Console.WriteLine($"JSON Reader Exception: {jsonReaderEx.Message}");
            }
            catch (JsonSerializationException jsonSerializationEx)
            {
                // Handle JSON serialization issues
                Console.WriteLine($"JSON Serialization Exception: {jsonSerializationEx.Message}");
            }
            catch (Exception ex)
            {
                // Handle any other unexpected errors
                Console.WriteLine($"An unexpected error occurred: {ex.Message}");
            }
        }
    }

    public static async Task<List<string?[]>> Translate_text(Dictionary<int, string> text, int chunkSize, int batchSize)
    {
        Dictionary<int, string> textChunk = new Dictionary<int, string>(); // contains a chunk of text (10)
        List<Task<string?>> batchTasks = new List<Task<string?>>(); // contains a list of chunks (5)
        List<string?[]> listOfTranslatedChunks = new List<string?[]>(); //contains every translated chunk

        int itemCount = 0;
        int currentBatch = 0;
        int totalChunks = (int)Math.Ceiling(text.Count / (double)chunkSize);

        //Math
        int fullBatchCount = totalChunks / batchSize;
        int partialBatch = (totalChunks % batchSize) > 0 ? 1 : 0;
        int maxBatch = fullBatchCount + partialBatch;

        foreach (var item in text)
        {
            itemCount++;
            textChunk.Add(item.Key, item.Value);

            if (textChunk.Count == chunkSize || itemCount == text.Count)
            {
                try
                {
                    string currentChunk = JsonConvert.SerializeObject(textChunk, Formatting.Indented);
                    Console.WriteLine(currentChunk);
                    batchTasks.Add(API(currentChunk));
                }
                catch (Exception ex) // Catch any exception that occurs within the try block
                {
                    Console.WriteLine($"An error occurred during serialization: {ex.Message}");
                    // Optionally, handle the error, e.g., by logging or setting a flag
                }
                finally
                {
                    textChunk.Clear(); // Ensure textChunk is cleared whether or not an exception occurs
                }
            }

            if (batchTasks.Count == batchSize || itemCount == text.Count)
            {
                currentBatch++;
                var batchResults = await Task.WhenAll(batchTasks); // translate
                listOfTranslatedChunks.Add(batchResults);
                Console.WriteLine($"Translating Batch: {currentBatch} / {maxBatch}");
                ProgressLog(batchResults);
                batchTasks.Clear();
            }

        }

        return listOfTranslatedChunks;
    }

    public static Dictionary<int, string> ParseExcelFile(string Path)
    {
        var fileContent = new FileInfo(Path);
        using var package = new ExcelPackage(fileContent);
        Dictionary<int, string> originalTextList = new Dictionary<int, string> { };

        // access WorkSheet
        ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
        var allRows = worksheet.Dimension.End.Row; // number of rows in the entire excel

        //Iterate through every row to get the rows in column 1
        for (int row = 1; row <= allRows; row++)
        {
            //add text 
            var value = worksheet.Cells[row, 1].Value?.ToString() ?? string.Empty;

            if (!IsJapanese(value))
            {
                continue;
            }

            originalTextList.Add(row, value);

        }

        return originalTextList;
    }

    public static void ImportTranslation(string path, string savePath, List<string?[]> translation)
    {
        var fileContent = new FileInfo(path);
        using var package = new ExcelPackage(fileContent);

        // access WorkSheet
        ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
        var allRows = worksheet.Dimension.End.Row; // number of rows in the entire excel

        // Iterate through each list of dictionaries
        foreach (var translationList in translation)
        {
            // Iterate through each dictionary in the list
            foreach (var json in translationList)
            {
                Dictionary<int, string> translationDict = null;
                try
                {
                    // Attempt to deserialize JSON outside the row loop
                     translationDict = JsonConvert.DeserializeObject<Dictionary<int, string>>(json);
                }
                catch (Exception ex) // Catch any exception that occurs during deserialization
                {
                    Console.WriteLine($"An error occurred during deserialization: {ex.Message}");
                    continue; // Skip to the next iteration in the loop
                }

                // Directly update cells based on dictionary keys without iterating through each row
                foreach (var item in translationDict)
                {
                    // Since item.Key is the row number, directly access the cell
                    if (item.Key >= 1 && item.Key <= allRows) // Check if the row exists within the range
                    {
                        worksheet.Cells[item.Key, 2].Value = item.Value;
                    }
                }
            }
        }

        // Save the package to a new file
        FileInfo newFile = new FileInfo(savePath);
         package.SaveAs(newFile);

    }

    public static async Task Main(string[] args)
    {
        string sourceDirectory = @"C:\Users\sanch\OneDrive\Desktop\Translation tools\temp_trans\Untranslated";
        string targetDirectory = @"C:\Users\sanch\OneDrive\Desktop\Translation tools\temp_trans\Translated";
        string[] excelFiles = Directory.GetFiles(sourceDirectory, "*.xlsx");
        int count = 0;
        int maxFile = excelFiles.Length;


         foreach (string filePath in excelFiles)
         {
             count++;
             Console.WriteLine($"Starting Translation {count} / {maxFile}");
             Console.WriteLine();

             // Extract the file name and create a new path for the translated file
             string fileName = Path.GetFileName(filePath);
             string newPath = Path.Combine(targetDirectory, fileName);

             // Operations to do
             Dictionary<int, string> extractedText = ParseExcelFile(filePath); // Export text from current Excel file
             List<string?[]> translatedText = await Translate_text(extractedText, 10, 2); // Translate text
             ImportTranslation(filePath, newPath, translatedText); // Import text back into an Excel file

             Console.WriteLine($"Completed Translation {count} / {maxFile}");
             Console.WriteLine();
         }
    }
}