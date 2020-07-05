using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Word;

namespace ReadWordDoc
{
    class Program
    {
        public static string path = "D:\\SampleProjects\\ReadWordDoc\\";
        public static Application word = new Microsoft.Office.Interop.Word.Application();
        public static Document doc = new Document();

        static void Main(string[] args)
        {
            Microsoft.Office.Interop.Word.Hyperlinks links1 = readDoc(path + "EmployeeDocument.docx");
            ReadHyperLinks(links1);
            ((_Document)doc).Close();
            ((_Application)word).Quit();
            Console.Read();
        }
        public static void ReadHyperLinks(Hyperlinks links1)
        {

            for (int i = 1; i <= links1.Count; i++)
            {
                object index = (object)i;
                string c = ((dynamic)links1[i]).Address;
                Hyperlinks links2 = readDoc(path + c);
                if (links2.Count > 0)
                {
                    ReadHyperLinks(links2);
                }
            }
        }
        public static Hyperlinks readDoc(string file)
        {

            object fileName = file;
            object missing = System.Type.Missing;
            doc = word.Documents.Open(ref fileName, false, ref missing, ref missing, ref missing,
                   ref missing, ref missing, ref missing, ref missing,
                   ref missing, ref missing, ref missing, ref missing,
                   ref missing, ref missing, ref missing);
            String read = string.Empty;
            Microsoft.Office.Interop.Word.Hyperlinks links = doc.Hyperlinks;
            List<string> data = new List<string>();
            foreach (Paragraph objParagraph in doc.Paragraphs)
            {
                Console.WriteLine(objParagraph.Range.Text.Trim());
            }

            Console.Write("******************** End Of File *******************************\n");

            return links;
        }


    }
}
