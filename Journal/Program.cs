using System;
using Word = Microsoft.Office.Interop.Word;

namespace Journal
{
    class Program
    {
        const string JOURNAL_PATH = @"C:\Users\engrb\OneDrive\bugs\Journal";
        static void Main(string[] args)
        {
            Word.Application objWord = new Word.Application();
            objWord.Visible = true;
            objWord.WindowState = Word.WdWindowState.wdWindowStateMaximize;

            Word.Document objDoc = objWord.Documents.Add();

            objWord.Selection.TypeText("Heading");
            
            objDoc.Paragraphs[1].set_Style(Word.WdBuiltinStyle.wdStyleHeading1);
            objDoc.Paragraphs[1].Range.Underline = Word.WdUnderline.wdUnderlineSingle;
            
            objWord.Selection.TypeParagraph();
            objWord.Selection.TypeText(Environment.NewLine);
            objDoc.Paragraphs[2].set_Style(Word.WdBuiltinStyle.wdStyleNormal);

            objDoc.SaveAs2( @"c:\temp\hello.docx");

            objWord.Activate();




            /*
            Word.Paragraph para1 = objDoc.Paragraphs.Add();
            object styleHeading1 = "Heading 1";
            para1.Range.set_Style(ref styleHeading1);

            para1.Range.Text = "Para 1 text";
            para1.Range.Underline = Word.WdUnderline.wdUnderlineSingle;
            para1.Range.InsertParagraphAfter();
            para1.Range.Select();
            

            Word.Paragraph para2 = objDoc.Paragraphs.Add();
            object styleNormal = "Normal";
            para2.Range.set_Style(ref styleNormal);
            // para2.Range.Underline = Word.WdUnderline.wdUnderlineNone;
            para2.Range.Text = Environment.NewLine;
            para2.Range.InsertParagraphAfter();
            */

            /*
            //  jump to the end of the document.
            object StartPos = 0;
            object Endpos = 1;
            Word.Range rng = objDoc.Range(ref StartPos, ref Endpos);

            object NewEndPos = rng.StoryLength - 1;
            Console.WriteLine(NewEndPos);
            rng = objDoc.Range(ref NewEndPos, ref NewEndPos);
            rng.Select();
            */





        }

        static string strHeader1Text()
        {
            return "Hello Word";

        }
    }
}
