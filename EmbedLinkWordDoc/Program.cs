using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Word;
using Microsoft.Office.Interop;
using Microsoft.Office;
using System.IO;

namespace EmbedLinkWordDoc
{
    class Program
    {
        private static Random randString = new Random();
        private static Random randLength = new Random();

        static void Main(string[] args)
        {
            try
            {
                CreateWordDocument(args[0]);
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                Console.WriteLine(e.StackTrace);
            }
        }

        static void CreateWordDocument(string Link)
        {
            try
            {
                Microsoft.Office.Interop.Word._Application MSWORD = new Microsoft.Office.Interop.Word.Application();
                MSWORD.Visible = false;
                object missing = Type.Missing;
                Microsoft.Office.Interop.Word._Document MSWORDdoc = MSWORD.Documents.Add(ref missing, ref missing, ref missing, ref missing);
                Microsoft.Office.Interop.Word.Paragraph TEXTBLOCK = MSWORDdoc.Paragraphs.Add(ref missing);
                object style_name = "Heading 1";
                TEXTBLOCK.Range.set_Style(ref style_name);
                TEXTBLOCK.Range.InsertParagraphAfter();
                TEXTBLOCK.Range.Text = RandomString(randLength.Next(5, 999));
                TEXTBLOCK.Range.InsertParagraphAfter();
                object filename = Directory.GetCurrentDirectory()+"\\"+ RandomString(randLength.Next(3,12))+".doc";
                string PREV_FONT = TEXTBLOCK.Range.Font.Name;
                TEXTBLOCK.Range.Font.Name = "Courier New";
                TEXTBLOCK.Range.InsertParagraphAfter();
                TEXTBLOCK.Range.Font.Name = PREV_FONT;
                Microsoft.Office.Interop.Word.Hyperlinks myLinks = MSWORDdoc.Hyperlinks;
                if (Link.Contains("http")==false && Link.Contains("https") == false)
                {
                    Link = "http://" + Link;
                }
                object linkAddr = Link;
                Microsoft.Office.Interop.Word.Selection mySelection = MSWORDdoc.ActiveWindow.Selection;
                mySelection.Start = 9999;
                mySelection.End = 9999;
                Microsoft.Office.Interop.Word.Range myRange = mySelection.Range;
                Microsoft.Office.Interop.Word.Hyperlink myLink = myLinks.Add(myRange, ref linkAddr, ref missing);
                MSWORDdoc.ActiveWindow.Selection.InsertAfter("\n");
                MSWORDdoc.SaveAs(ref filename, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,ref missing);
                object save_changes = false;
                MSWORDdoc.Close(ref save_changes, ref missing, ref missing);
                MSWORD.Quit(ref save_changes, ref missing, ref missing);
                Console.WriteLine("App done! .Doc created!");
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                Console.WriteLine(e.StackTrace);
            }
        }

        private static string RandomString(int length)
        {
            const string chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789      ";
            return new string(Enumerable.Repeat(chars, length)
              .Select(s => s[randString.Next(s.Length)]).ToArray());
        }
    }
}
