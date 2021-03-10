using System;
using Word = Microsoft.Office.Interop.Word;
using System.IO;


namespace Journal
{
    class Program
    {
        const string VERSION = "v0.91";
        const string JOURNAL_PATH = @"C:\Users\engrb\OneDrive\bugs\Journal";
        static void Main(string[] args)
        {
            Console.WriteLine($"Creating Journal {VERSION}");
            Console.WriteLine($"Loading WinWord App");
            Word.Application objWord = new Word.Application();
            Console.WriteLine($"Making it visible and maximized");
            objWord.Visible = true;
            objWord.WindowState = Word.WdWindowState.wdWindowStateMaximize;

            Console.WriteLine($"Creating a document");
            Word.Document objDoc = objWord.Documents.Add();

            Console.WriteLine($"What is the date for today... hmmm...");
            objWord.Selection.TypeText(strHeader1Text());

            Console.WriteLine($"Typing the header");
            objDoc.Paragraphs[1].set_Style(Word.WdBuiltinStyle.wdStyleHeading1);
            objDoc.Paragraphs[1].Range.Underline = Word.WdUnderline.wdUnderlineSingle;

            Console.WriteLine($"Clear formatting for you to type");
            objWord.Selection.TypeParagraph();
            objWord.Selection.TypeText(Environment.NewLine);
            objDoc.Paragraphs[2].set_Style(Word.WdBuiltinStyle.wdStyleNormal);

            try
            {
                Console.WriteLine($"Making a folder");
                Directory.CreateDirectory($"{JOURNAL_PATH}\\" +
                    $"{strFolderYear()}\\{strFolderMonth()}");
                Console.WriteLine($"I'm close... Saving the file for you");
                objDoc.SaveAs2($"{JOURNAL_PATH}\\{strFolderYear()}\\" +
                    $"{strFolderMonth()}\\{strFilename()}.docx");
            }
            catch (Exception e)
            {
                Console.WriteLine("The process failed: {0}", e.ToString());
            }
            
            Console.WriteLine($"Focus on WinWord app");
            objWord.Activate();

            Console.WriteLine($"Bye.");
        }

        static string strHeader1Text()
        {
            return $"{DateTime.Now.DayOfWeek}, " +
                $"{DateTime.Now.ToLongDateString()}" +
                $"—{DateTime.Now.ToShortTimeString()}";
        }

        static string strFolderYear()
        {
            return DateTime.Now.Year.ToString();
        }

        static string strFolderMonth()
        {
            return $"{DateTime.Now.Month.ToString("00")}-" +
                $"{DateTime.Now.ToString("MMMM")}";
        }

        static string strFilename()
        {
            return $"{DateTime.Now.ToString("d")}-" +
                $"{DateTime.Now.ToString("HH")}" +
                $"{DateTime.Now.ToString("mm")}";
        }
    }
}
