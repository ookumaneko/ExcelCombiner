using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;

using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;

namespace ExcelCombiner
{
    class Converter
    {
        const string _SHEET_NAME_NO_NUMBER = "ScenarioData";
        const string _SHEET_NAME = _SHEET_NAME_NO_NUMBER + "{0:D4}";

        const int _START_ROW = 1;

        string m_directory;
        string m_outputDirectory;

        public Converter(string directory, string outputDirectory)
        {
            m_directory = directory;
            m_outputDirectory = outputDirectory;
        }

        public void Convert(string[] originalFiles)
        {
            try
            {
                Logger.WriteLine(Environment.NewLine + "Creating new book at: " + m_outputDirectory + Environment.NewLine);
                var book = NPOIUtility.CreateNewBook(m_outputDirectory);

                int fileCount = originalFiles.Length;
                Logger.WriteLine("Start Processing...{0} files", fileCount);
                for (int i = 0; i < fileCount; ++i)
                {
                    Logger.WriteLine("...processing file [{0}/{1}]", i + 1, fileCount);
                    CopySheet(book as HSSFWorkbook, i, originalFiles[i]);
                }

                using (var fs = File.Create(m_outputDirectory))
                {
                    book.Write(fs);
                    book.Close();
                }
            }
            catch (Exception ex)
            {
                Logger.WriteLine(ex.Message);
            }
        }

        private void CopySheet(HSSFWorkbook book, int index, string loadFilePath)
        {
            IWorkbook loadedBook = NPOIUtility.LoadBook(loadFilePath);
            if (loadedBook == null)
            {
                return;
            }

            ((HSSFSheet)loadedBook.GetSheetAt(0)).CopyTo(book, loadedBook.GetSheetName(0), true, true);
        }
    }
}
