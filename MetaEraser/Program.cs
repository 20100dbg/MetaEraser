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

        static List<String> supportedExt = new List<String> { ".docx", ".docm", ".xlsx", ".xlsm", ".pptx", ".pptm" };
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

            Console.WriteLine("MetaEraser va traiter les fichiers dans le dossier in");
            Console.WriteLine("L'original sera conservé dans le dossier in, les fichiers traités sont dans le dossier out");
            Console.WriteLine("Tapez exit ou quit pour ne pas traiter les fichiers");
            String input = Console.ReadLine();

            if (input.ToLower() != "exit"  && input.ToLower() != "quit")
                cleanAll();


            Console.WriteLine("[+] Finished. Everything is in " + pathOut);
            Console.Read();
        }

        static void cleanAll()
        {
            String[] files = Directory.GetFiles(pathIn);

            for (int i = 0, n = files.Length; i < n; i++)
            {
                String filename = Path.GetFileName(files[i]);
                String ext = Path.GetExtension(files[i]).ToLower();

                if (supportedExt.Contains(ext)) cleanOOXML(pathIn, filename);
                else if (ext == ".pdf") cleanPdf(pathIn, filename);
                else
                {
                    System.IO.File.Copy(pathIn + filename, pathOut + filename, true);
                    //Console.WriteLine("[-] Ignore " + filename);
                }
                Console.WriteLine("[+] Process " + filename);
            }

            System.Diagnostics.Process.Start(Directory.GetCurrentDirectory() + "\\exiftool.exe", "-overwrite_original -all= out/*.*");
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