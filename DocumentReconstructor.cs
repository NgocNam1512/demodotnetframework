using System;
using System.Drawing;
using System.IO;
using Word = Microsoft.Office.Interop.Word;
using Microsoft.Office.Core;
using System.Reflection;
using System.Text.Json;

namespace Reconstructor
{
    public class DocumentReconstructor
    {
        public static string imgToBase64(Image img)
        {
            using (MemoryStream m = new MemoryStream())
            {
                img.Save(m, img.RawFormat);
                byte[] imageBytes = m.ToArray();

                // Convert byte[] to Base64 String
                string base64String = Convert.ToBase64String(imageBytes);
                return base64String;
            }
        }
        public static Image base64ToImage(string base64String)
        {
            // Convert base 64 string to byte[]
            byte[] imageBytes = Convert.FromBase64String(base64String);
            // Convert byte[] to Image
            using (var ms = new MemoryStream(imageBytes, 0, imageBytes.Length))
            {
                Image image = Image.FromStream(ms, true);
                return image;
            }
        }
        //convert image to bytearray
        public static byte[] imgToByteArray(Image img)
        {
            using (MemoryStream mStream = new MemoryStream())
            {
                img.Save(mStream, img.RawFormat);
                return mStream.ToArray();
            }
        }
        //convert bytearray to image
        public static Image byteArrayToImage(byte[] byteArrayIn)
        {
            using (MemoryStream mStream = new MemoryStream(byteArrayIn))
            {
                return Image.FromStream(mStream);
            }
        }
        public static void InsertParagraph(Word.Document doc, string content)
        {
            Word.Paragraph para;
            para = doc.Content.Paragraphs.Add();
            para.Range.Text = content;
            para.Range.Font.Name = "Time New Romans";
            para.Range.Font.Size = 11;
            //para.Range.Font.Bold = 1;
            para.Range.InsertParagraphAfter();
            para.Format.SpaceAfter = 6;
            para.Format.FirstLineIndent = 1;
        }
        public static void InsertTextbox(Word.Document doc, Int32 paper_width, Int32 paper_height, string content, Location location)
        {
            float A4_point_width = 595;
            float A4_point_height = 842;

            int left = (int)((float)location.x1 / (float)paper_width * A4_point_width);
            int top = (int)((float)location.y1 / (float)paper_height * A4_point_height);
            int width = (int)(((float)location.x2 - (float)location.x1) / (float)paper_width * A4_point_width);
            int height = (int)(((float)location.y2 - (float)location.y1) / (float)paper_height * A4_point_height * 1.5);

            Word.Shape textbox;
            textbox = doc.Shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal, left, top, width, height);
            textbox.TextFrame.TextRange.Text = content;
            textbox.TextFrame.TextRange.Font.Name = "Time New Romans";
            textbox.TextFrame.TextRange.Font.Size = 13;
            textbox.Line.Visible = MsoTriState.msoFalse;
        }
        public static void InsertImage(Word.Document doc, Int32 paper_width, Int32 paper_height, string base64Image, Location location)
        {
            float A4_point_width = 500;
            float A4_point_height = 705;
            string imageName = "temp.jpg";
            Image img = base64ToImage(base64Image);
            img.Save(imageName);

            int left = (int)((float)location.x1 / (float)paper_width * A4_point_width);
            int top = (int)((float)location.y1 / (float)paper_height * A4_point_height);
            int width = (int)(((float)location.x2 - (float)location.x1) / (float)paper_width * A4_point_width);
            int height = (int)(((float)location.y2 - (float)location.y1) / (float)paper_height * A4_point_height);

            Word.Shape image;
            string imagepath = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "temp.jpg");
            image = doc.Shapes.AddPicture(imagepath, false, true, left, top, width, height);
            image.WrapFormat.AllowOverlap = 0;
            image.WrapFormat.Type = Word.WdWrapType.wdWrapThrough;
        }
        public static bool CreateDocument(string jsonpath, string output_path)
        {

            Data data;
            using (StreamReader r = new StreamReader(jsonpath))
            {
                string json = r.ReadToEnd();
                Console.WriteLine(json);
                data = JsonSerializer.Deserialize<Data>(json);
            }
            int width = data.width;
            int height = (int)(width * 1.41);
            ObjData[] datalist = data.datalist;

            object missing = System.Reflection.Missing.Value;
            object endOfDoc = "\\endofdoc";
            Word.Application app = new Word.Application();
            Word.Document document;
            app.Visible = true;
            document = app.Documents.Add();

            for (int i = 0; i < datalist.Length; i++)
            {
                ObjData obj = datalist[i];
                if (obj.label == "line")
                {
                    //InsertParagraph(document, obj.content);
                    InsertTextbox(document, width, height, obj.content, obj.location);
                }
                else if (obj.label == "textbox")
                {
                    InsertTextbox(document, width, height, obj.content, obj.location);
                }
                else if (obj.label == "image")
                {
                    InsertImage(document, width, height, obj.content, obj.location);
                }
                else
                {
                    Console.WriteLine("Don't have this label!!!");
                }
            }

            // save document
            object filename = output_path;
            document.SaveAs2(ref filename);
            document.Close();
            document = null;
            app.Quit();
            app = null;
            Console.WriteLine("Document created successfully !");
            return true;
        }
    }
}
