namespace DocxConnector
{
    using System;
    using System.IO;
    using System.Linq;
    using Word = Microsoft.Office.Interop.Word;

    class Program
    {
        public static string nameResult = Directory.GetCurrentDirectory() + "\\result.docx";
        static void Main(string[] args)
        {
            if (!string.IsNullOrEmpty(Directory.GetFiles(Directory.GetCurrentDirectory(), "*.docx").FirstOrDefault(u => u == nameResult)))
            {
                File.Delete(nameResult);
            }

            var docxDocuments = Directory.GetFiles(Directory.GetCurrentDirectory(), "*.docx").ToArray();
            Merge(docxDocuments, nameResult);
        }
        public static void Merge(string[] filesToMerge, string outputFilename)
        {
            object missing = Type.Missing;
            object pageBreak = Word.WdBreakType.wdPageBreak;
            object outputFile = outputFilename;
            Word._Application wordApplication = new Word.Application();
            try
            {
                Word._Document wordDocument = wordApplication.Documents.Add(
                                              ref missing
                                            , ref missing
                                            , ref missing
                                            , ref missing);
                Word.Selection selection = wordApplication.Selection;
                foreach (string file in filesToMerge)
                {
                    selection.InsertFile(
                                            file
                                            , ref missing
                                            , ref missing
                                            , ref missing
                                            , ref missing);
                }
                wordDocument.SaveAs(
                            ref outputFile
                            , ref missing
                            , ref missing
                            , ref missing
                            , ref missing
                            , ref missing
                            , ref missing
                            , ref missing
                            , ref missing
                            , ref missing
                            , ref missing
                            , ref missing
                            , ref missing
                            , ref missing
                            , ref missing
                            , ref missing);
                wordDocument = null;
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                wordApplication.Quit(ref missing, ref missing, ref missing);
            }
        }
    }
}