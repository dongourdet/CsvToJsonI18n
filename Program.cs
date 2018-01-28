﻿using System;
using System.IO;
using OfficeOpenXml;
using System.Linq;
using System.Collections.Generic;
using Newtonsoft.Json.Linq;
using System.Text;

namespace CsvToJson
{
    class Program
    {
        static void Main(string[] args)
        {
            FileInfo fileInfo = new FileInfo(@"c:\temp\test.xlsx");
            var languages = ReadFile(fileInfo);
            JObject jObject;
            languages.Skip(1).ToList().ForEach(language =>
            {
                jObject = new JObject();
                language.Translations.ForEach(translation =>
                {
                    jObject[translation.Key] = translation.Value;
                });
                WriteFile(jObject, fileInfo, language.Key);                
            });            
            
            Console.WriteLine("Press Enter to exit!");
            Console.ReadLine();
        }
        private static List<Language> ReadFile(FileInfo fileInfo)
        {
            using (var package = new ExcelPackage(fileInfo))
            {
                var workbook = package.Workbook;
                var cells = workbook.Worksheets.First().Cells.ToList();
                var headers = cells.GroupBy(x => x.Start.Column).ToList();

                var columns = cells.GroupBy(x => x.Start.Column).ToList();

                var languages = columns.Select((hcell, colIndex) => new Language
                {
                    Row = 1,
                    Col = hcell.Select(x => x.Start.Column).FirstOrDefault(),
                    Key = hcell.Select(x => x.Value).FirstOrDefault().ToString(),

                    Translations = hcell.Skip(1).Select((cell, rowIndex) => new Translation
                    {
                        LangKey = hcell.Select(x => x.Value).FirstOrDefault().ToString(),
                        Key = columns.FirstOrDefault().Where(x=>x.Start.Row == cell.Start.Row).Select(x=>x.Value).FirstOrDefault().ToString(),
                        Value = cell.Value.ToString(),
                        Row = cell.Start.Row,
                        Col = cell.Start.Column
                    }).ToList()
                }).ToList();
                return languages;
            }
        }
        private static void WriteFile(JObject contents, FileInfo fileInfo, string languageKey)
        {
            var fileName = string.Format("{0}.json", languageKey);
            var filePath = string.Format("{0}\\{1}", fileInfo.DirectoryName, fileName);
            var fi = new FileInfo(filePath);
            if (fi.Exists)
            {
                fi.Delete();
            }
            using (var fs = fi.Create())
            {
                Byte[] bytes = new UTF8Encoding(true).GetBytes(contents.ToString());
                fs.Write(bytes, 0, bytes.Length);
            }
        }
    }
    class Language
    {
        public string Key { get; set; }
        public List<Translation> Translations { get; set; }
        public int Col { get; set; }
        public int Row { get; set; }

        public Language()
        {
        }


    }
    class Translation
    {
        public string Key { get; set; }
        public string LangKey { get; set; }
        public string Value { get; set; }
        public int Col { get; set; }
        public int Row { get; set; }
        public Translation()
        {
        }
    }


}