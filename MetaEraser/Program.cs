using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using System.Xml;
using System.Text;

using Ionic.Zip;
using org.pdfclown.documents;
using org.pdfclown.files;
using org.pdfclown.documents.interchange.metadata;


namespace MetaEraser
{
    class Program
    {

        static List<String> supportedExt = new List<String> { ".pptx", ".docx", ".xlsx", ".xlsm" };
        static String pathIn = "in/";
        static String pathOut = "out/";

        static void Main(string[] args)
        {

            if (!Directory.Exists(pathIn))
            {
                Console.WriteLine(pathIn + " introuvable");
                Console.Read();
                return;
            }

            if (!Directory.Exists(pathOut))
            {
                Console.WriteLine(pathOut + " introuvable");
                Console.Read();
                return;
            }


            String[] files = Directory.GetFiles(pathIn);

            for (int i = 0, n = files.Length; i < n; i++)
            {
                String filename = Path.GetFileName(files[i]);
                String ext = Path.GetExtension(files[i]).ToLower();

                if (supportedExt.Contains(ext))
                {
                    Console.WriteLine("[+] Process " + filename);
                    cleanOOXML(pathIn, filename);
                }
                else if (ext == ".pdf")
                {
                    Console.WriteLine("[+] Process " + filename);
                    cleanPdf(pathIn, filename);
                }
                else
                {
                    Console.WriteLine("[-] Ignore " + filename);
                }
            }

            Console.WriteLine("[+] Finished. Everything is in " + pathOut);
            Console.Read();
        }



        static void cleanPdf(String path, String filename)
        {
            using (org.pdfclown.files.File file = new org.pdfclown.files.File(path + filename))
            {
                Document document = file.Document;
                Information info = document.Information;
                if (info.Exists())
                {
                    document.Information.Clear();
                }

                Metadata metadata = document.Metadata;
                if (metadata.Exists())
                {
                    metadata.Content.RemoveAll();
                }

                document.File.Save(pathOut + filename, SerializationModeEnum.Standard);
            }

        }



        static void cleanOOXML(String path, String filename)
        {
            XmlDocument xmldoc = new XmlDocument();
            MemoryStream ms = new MemoryStream();
            String str = "";

            List<String> tagsApp = new List<String> { "Manager", "Company", "Template" }; 
            List<String> tagsCore = new List<String> { "dc:title", "dc:subject", "dc:creator","cp:keywords", "dc:description","cp:lastModifiedBy", "cp:category","cp:contentStatus", "dc:language","cp:version" };

            Regex rgx;

            using (ZipFile zip = ZipFile.Read(path + filename))
            {

                ZipEntry e = zip["docProps/core.xml"];
                e.Extract(ms);

                Byte[] buffer = new Byte[ms.Length];
                ms.Seek(0, SeekOrigin.Begin);
                ms.Read(buffer, 0, buffer.Length);
                str = UTF8Encoding.UTF8.GetString(buffer);

                for (int i = 0; i < tagsCore.Count;i++)
                {
                    rgx = new Regex("<" + tagsCore[i] + ">" + "(.+?)" + "</" + tagsCore[i] + ">");
                    Match m = rgx.Match(str);
                    if (m.Success)
                        str = str.Replace(m.Value, "<" + tagsCore[i] + "></" + tagsCore[i] + ">");
                }

                zip.UpdateEntry("docProps/core.xml", str);
                
                e = zip["docProps/app.xml"];
                ms.SetLength(0);
                e.Extract(ms);

                buffer = new Byte[ms.Length];
                ms.Seek(0, SeekOrigin.Begin);
                ms.Read(buffer, 0, buffer.Length);
                str = UTF8Encoding.UTF8.GetString(buffer);

                for (int i = 0; i < tagsApp.Count; i++)
                {
                    rgx = new Regex("<" + tagsApp[i] + ">" + "(.+?)" + "</" + tagsApp[i] + ">");
                    Match m = rgx.Match(str);
                    if (m.Success)
                        str = str.Replace(m.Value, "<" + tagsApp[i] + "></" + tagsApp[i] + ">");
                }

                
                zip.UpdateEntry("docProps/app.xml", str, Encoding.ASCII); //l'encoding utf8 ne passe pas
                
                zip.Comment = "";
                zip.Save(pathOut + filename);

            }

        }


    }
}