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
        private static Random randString = new Random(DateTime.Now.Millisecond);
        private static Random randLength = new Random(DateTime.Now.Millisecond);
        private static Random rand_Word = new Random(DateTime.Now.Millisecond);
        private static string FilePath = "";

        static void Main(string[] args)
        {
            try
            {
                if (args.Length == 1)
                {
                    CreateWordDocument(args[0]);
                }
                else
                {
                    CreateWordDocument(args[0],args[1]);
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                Console.WriteLine(e.StackTrace);
            }
        }

        private static void CreateWordDocument(string Link,string FilePath="")
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
                if (string.IsNullOrEmpty(FilePath))
                { 
                    TEXTBLOCK.Range.Text = RandomString_Contents(randLength.Next(5, 999));
                }
                else
                {
                    TEXTBLOCK.Range.Text = RandomString_Words(FilePath);
                }
                TEXTBLOCK.Range.InsertParagraphAfter();
                object filename = Directory.GetCurrentDirectory() + "\\" + RandomString_Contents(randLength.Next(3, 12)) + ".doc";
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

        private static string RandomString_Contents(int length)
        {
            const string chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789      ";
            return new string(Enumerable.Repeat(chars, length).Select(s => s[randString.Next(s.Length)]).ToArray());
        }

        private static string RandomString_Words(string FilePath)
        {
            string[] FileContents = File.ReadAllLines(FilePath);
            return new string(Enumerable.Repeat(FileContents[rand_Word.Next(0, FileContents.Length-1)], FileContents.Length - 1).Select(s => s[randString.Next(s.Length)]).ToArray());
        }
    }
}
