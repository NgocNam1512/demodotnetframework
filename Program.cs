using System;
using System.Drawing;
using System.IO;
using Word = Microsoft.Office.Interop.Word;
using Microsoft.Office.Core;
using System.Reflection;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace demodotnetframework
{
    public class Location
    {
        public int x1 { get; set; }
        public int y1 { get; set; }
        public int x2 { get; set; }
        public int y2 { get; set; }
    }
    public class ObjData
    {
        public string label { get; set; }
        public Location location { get; set; }
        public string content { get; set; }
    }
    public class Data
    {
        public int width { get; set; }
        public ObjData[] datalist { get; set; }
    }
    class Program
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
            para.Range.Font.Size = 13;
            //para.Range.Font.Bold = 1;
            para.Range.InsertParagraphAfter();
            para.Format.SpaceAfter = 24;
            para.Format.FirstLineIndent = 1;
        }
        public static void InsertTextbox(Word.Document doc, string content, Location location)
        {
            Int32 left = location.x1;
            Int32 top = location.y1;
            Int32 width = location.x2 - location.x1;
            Int32 height = location.y2 - location.y1;

            Word.Shape textbox;
            textbox = doc.Shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal, left, top, width, height);
            textbox.TextFrame.TextRange.Text = content;
        }
        public static void InsertImage(Word.Document doc, string base64Image, Location location)
        {
            string imageName = "temp.jpg";
            Image img = base64ToImage(base64Image);
            img.Save(imageName);

            Int32 left = location.x1;
            Int32 top = location.y1;
            Int32 width = location.x2 - location.x1;
            Int32 height = location.y2 - location.y1;

            Word.Shape image;
            string imagepath = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "2.jpg");
            image = doc.Shapes.AddPicture(imagepath, false, true, left, top, width, height);
            image.WrapFormat.AllowOverlap = 0;
            //image = doc.Shapes.AddShape(1, left, top, width, height);
            //image.Fill.UserPicture(
            //    Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "2.jpg"));
        }
        public static void CreateDocument(Data data)
        {
            int width = data.width;
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
                if (obj.label == "text")
                {
                    InsertParagraph(document, obj.content);
                }
                else if (obj.label == "textbox")
                {
                    InsertImage(document, obj.content, obj.location);
                }
                else if (obj.label == "image")
                {
                    InsertImage(document, obj.content, obj.location);
                }
                else
                {
                    Console.WriteLine("Done have this label!!!");
                }
            }

            //string content = "Lorem Ipsum is simply dummy text of the printing and typesetting industry. Lorem Ipsum has been the industry's standard dummy text ever since the 1500s, when an unknown printer took a galley of type and scrambled it to make a type specimen book. It has survived not only five centuries, but also the leap into electronic typesetting, remaining essentially unchanged. It was popularised in the 1960s with the release of Letraset sheets containing Lorem Ipsum passages, and more recently with desktop publishing software like Aldus PageMaker including versions of Lorem Ipsum.";
            //InsertParagraph(document, content);
            //InsertParagraph(document, width.ToString());
            //InsertTextbox(document, "this is a textbox", 200, 200, 100, 15);
            //var base64SignatureImage = "iVBORw0KGgoAAAANSUhEUgAAAOgAAACJCAMAAAAR34S1AAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAMAUExURQAAAA8PDx8fHy8vLz8/P09PT19fX29vb39/f4+Pj5+fn6+vr7+/v8/Pz9/f3+/v7////wAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAK/M6GAAAAEAdFJOU////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////wBT9wclAAAACXBIWXMAAA7DAAAOwwHHb6hkAAAAGHRFWHRTb2Z0d2FyZQBwYWludC5uZXQgNC4wLjb8jGPfAAALWklEQVR4Xu2da6Ojqg6GQVFREdb//7PdCSAXwUsLtp211/thH6fTaX0ghCTEHvL4n+gP9LfpD/S36Q/0t+kP9LfpDzSvhXcNQVHGF/vav6FnQFWPkJRpUUKYsH/xL+g6qGBAOUzK/vEhx4b09vof0FVQxSlpPaXRRP8d0ougM2La60BLM9irezWVL5JLoKoDo7XXsRb6DpfUk8Zeva4roEtL6Gyvtxo7e3GjJnAOxVN6AXShhG0Wp5cg0l7dJ0rGoXiJnIMC54HLEeVjfaYJlo1s7R9e1imoojvL02ig9uI+tRTsqXiRnoIyMtqrrJrbNxipB5qVOr0zUE6YvcqK379EzVfw0hVyAqoo2s2uFsrt1X1q9fK8G3Qke/sKaqb3by7KLB2eCVee0gloc2S4I2mPpruOJrM4eKnpHIPOBxMqWtLdz/nojL/lR5Z1Rcegw94KVTNEhaXWdEnUbG43r1GWXYNqBErC7w+JQGtAcjMoyayMpYOcm3Rvqi9wYmxquBk0CRYUFhna4W1lFGa94X60fVEnoKUuoFirTd0cAr4hHjiWsEMtbwZlny6VjNYXjcU3cgz6htzkWL0BVeWmdQy6fGqRKjFzrouqLdZWm8NU8ZKOQR9NebHmWcmZAx9lHDSYKjLuZ33hkJ+AYnb/Rqm5B6hmGG0wMtqv74G0MBA7AX00x3l3VWFYCZTB1HXEbNgNLNG5KYoZzkCXlrTviQ4gSSBNH8PYlbPokFsVed4zUAzCSH9/WCthNtl2GZoqCliu8bmixHjPQfEebg8cloZkLHMylquoHegSf3EBVFtVc+s+s1Nq7E2Rc1pzqBLbvQSKnuDW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";
            //InsertImage(document, base64SignatureImage, 0, 0, 100, 100);

            // save document
            object filename = @"C:\Users\namnn12\Desktop\demo2.docx";
            document.SaveAs2(ref filename);
            document.Close();
            document = null;
            app.Quit();
            app = null;
            Console.WriteLine("Document created successfully !");
        }
        static void Main(string[] args)
        {
            string jsonpath = @"C:\Users\namnn12\Desktop\testjson\2.json";

            Data data;
            using (StreamReader r = new StreamReader(jsonpath))
            {
                string json = r.ReadToEnd();
                data = JsonSerializer.Deserialize<Data>(json);
                CreateDocument(data);
            }
            Console.WriteLine(data.width);
        }
    }
}
