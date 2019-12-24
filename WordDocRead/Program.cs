using System;
using System.Collections.Generic;
using System.Linq;
using System.Data;
using System.Data.SqlClient;
using System.Runtime.InteropServices.ComTypes;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Word;
using System.Collections;
using System.Text.RegularExpressions;
using MySql.Data.MySqlClient;
/*
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;
*/
using System.Drawing;
using System.IO;

namespace WordDocRead
{
    class Program
    {
        static MySqlConnection connection;
        

        /*
        private static void ExtractImages(List<string> imagepaths)
        {
            //Load document  
            Spire.Doc.Document document = new Spire.Doc.Document(@"D:\TEMPORARY\Documents\temp.docx");
            int index = 0;

            //Get Each Section of Document  
            foreach (Spire.Doc.Section section in document.Sections)
            {
                //Get Each Paragraph of Section  
                foreach (Spire.Doc.Documents.Paragraph paragraph in section.Paragraphs)
                {
                    //Get Each Document Object of Paragraph Items  
                    foreach (DocumentObject docObject in paragraph.ChildObjects)
                    {
                        //If Type of Document Object is Picture, Extract.  
                        if (docObject.DocumentObjectType == DocumentObjectType.Picture)
                        {
                            //Console.WriteLine("image found");
                            DocPicture pic = docObject as DocPicture;
                            String imgName = String.Format(@"D:\TEMPORARY\Extracted_Image-{0}.png", index);
                            imagepaths.Add(imgName);
                            //Save Image  
                            pic.Image.Save(imgName, System.Drawing.Imaging.ImageFormat.Png);
                            index++;
                        }
                    }
                }
            }
        }
        */

        static void Main(string[] args)
        {
            List<string> imagepaths = new List<string>();
            //connectDB();
            //ExtractImages(imagepaths);
            //StringBuilder text = new StringBuilder();
            List<string> text = new List<string>();
            Microsoft.Office.Interop.Word.Application word = new Microsoft.Office.Interop.Word.Application();
            object miss = System.Reflection.Missing.Value;
            object path = @"D:\TEMPORARY\Documents\Hello22.docx";
            //object readOnly = true;
            Microsoft.Office.Interop.Word.Document docs = word.Documents.Open(ref path, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss);

            //Console.WriteLine(docs.Paragraphs.Count);


           

            for (int i = 0; i < docs.Paragraphs.Count; i++)
            {
                string s = docs.Paragraphs[i + 1].Range.Text.ToString();


                s = s.TrimEnd('\r');
                text.Add(s);
            }



            //Console.WriteLine(docs.Paragraphs.Count);
            
            for(int i = 0; i < text.Count(); i++)
            {
                
                Regex regex = new Regex(@"\d\)");
                Match match = regex.Match(text[i]);
                //Console.WriteLine(text[i]);
                if (text[i].Length == 2 && match.Success)
                {
                    Regex rgx2 = new Regex("\\)");
                    int qno = Int32.Parse(rgx2.Replace(text[i], ""));
                    //Console.WriteLine("MMMMMMMMMMMMMMM");
                    newQuestion(text, i,qno);
                }
            }
            //Console.WriteLine(text.ToString());

            docs.Close();
            //File.Replace((string)path, (string)path, @"D:\TEMPORARY");
            Console.ReadLine();
            
            //File.Replace((string)path, (string)path , @"D:\TEMPORARY");
        }

        static void newQuestion(List<string> text, int i, int qno)
        {
            string type;
            if (i + 1 != text.Count)
            {
                Regex rgx = new Regex("Type=");
                type = rgx.Replace(text[i + 1], "");
                //type.TrimStart('\t');
                
                if (type.Equals("mcq1"))
                {
                    mcq1(text, i + 1, qno);
                }
                
            }
        }

        static void mcq1(List<string> text, int i, int qno)
        {
            string question = "";
            List<string> options = new List<string>();
            List<string> answers = new List<string>();
            if (text[i + 1].Equals("Question="))
            {
                i++;
                Regex rgx = new Regex("Image=");
                string imageBoolean = rgx.Replace(text[i + 1], "");
                i++;
                if (imageBoolean.Equals("no"))
                {
                    while (!text[i + 1].Equals("QuestionEnd"))
                    {
                        question += text[i + 1] + " ";
                        i++;
                    }


                    //QUESTION QUERY HERE
                    
                    string connectionString = @"server=localhost;userid=n7aakash;password=1234@abcd;database=testdb";
                    string queryString = "INSERT into question(question_no,question_text) values(" + qno + ",'" + question + "')";
                    //string queryString = "SELECT OrderID, CustomerID FROM dbo.Orders;";
                    using (MySqlConnection connection = new MySqlConnection(connectionString))
                    {
                        MySqlCommand command = new MySqlCommand(queryString, connection);  //<- See here the connection is passes to the command
                        connection.Open();
                        MySqlDataReader reader = command.ExecuteReader();
                        try
                        {
                            
                        }
                        finally
                        {
                            // Always call Close when done reading.
                            reader.Close();
                        }
                    }

                    Console.WriteLine(question);
                    i++;
                    if (text[i + 1].Equals("Option="))
                    {
                        i++;
                        Regex rgx2 = new Regex("Image=");
                        string imageBoolean2 = rgx2.Replace(text[i + 1], "");
                        i++;
                        if (imageBoolean2.Equals("no"))
                        {
                            i = addOptions(text, i , options);


                            //OPTION QUERY HERE
                            foreach(string optn in options)
                            {
                                Console.WriteLine(optn);
                            }
                            if (text[i + 1].Equals("Answer="))
                            {
                                i=i+2;
                                //Console.WriteLine(text[i]);
                                
                                while(i!=text.Count && !text[i].Equals(""))
                                {
                                    //Console.WriteLine(text[i]);
                                    answers.Add(text[i]);
                                    i++;
                                }


                                //ANSWER QUERY HERE
                                foreach(string ans in answers)
                                {
                                    Console.WriteLine(ans);
                                }
                                
                            }
                        }
                        else
                        {
                            //code for image in options
                        }

                    }

                }
                else
                {
                    //code for image in question
                }
            }
            //Console.WriteLine(question);
            question = "";
            options.Clear();
            answers.Clear();
            Console.WriteLine("========================================");
        }

        

        static int addOptions(List<string> text, int i , List<string> options)
        {
            
            Regex rgx = new Regex("^[A-Z]=$");
            string optn = "";
            for(int j=i+1;j<text.Count;j++)
            {
                if (text[j].Equals("OptionEnd"))
                {
                    return j;
                }
                Match match = rgx.Match(text[j]);
                if (match.Success)
                {
                    
                    int temp = j+1;
                    //Console.WriteLine("Temp is" + text[temp]);
                    
                    while (!rgx.Match(text[temp]).Success && !text[temp].Equals("OptionEnd"))
                    {
                        optn = optn + text[temp] + " ";
                        temp++;
                    }
                    //Console.WriteLine("Option formed is :" + optn);
                    options.Add(optn);
                    optn = "";
                    //Console.WriteLine(text[j]);
                }
                
            }
            return i;
            
        }
    }
}
