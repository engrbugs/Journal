using System;
using Word = Microsoft.Office.Interop.Word;
using System.IO;


namespace Journal
{
    class Program
    {
        const string VERSION = "b09";
        const string JOURNAL_PATH = @"C:\Users\engrb\OneDrive\bugs\Journal";
        static void Main(string[] args)
        {
            Console.WriteLine($"Creating Journal {VERSION}");

            Word.Application objWord = new Word.Application();
            objWord.Visible = true;
            objWord.WindowState = Word.WdWindowState.wdWindowStateMaximize;

            Word.Document objDoc = objWord.Documents.Add();

            objWord.Selection.TypeText(strHeader1Text());
            
            objDoc.Paragraphs[1].set_Style(Word.WdBuiltinStyle.wdStyleHeading1);
            objDoc.Paragraphs[1].Range.Underline = Word.WdUnderline.wdUnderlineSingle;
            
            objWord.Selection.TypeParagraph();
            objWord.Selection.TypeText(Environment.NewLine);
            objDoc.Paragraphs[2].set_Style(Word.WdBuiltinStyle.wdStyleNormal);

            try
            {
                Directory.CreateDirectory($"{JOURNAL_PATH}\\" +
                    $"{strFolderYear()}\\{strFolderMonth()}");
                objDoc.SaveAs2($"{JOURNAL_PATH}\\{strFolderYear()}\\" +
                    $"{strFolderMonth()}\\{strFilename()}.docx");
            }
            catch (Exception e)
            {
                Console.WriteLine("The process failed: {0}", e.ToString());
            }
            objWord.Activate();
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
