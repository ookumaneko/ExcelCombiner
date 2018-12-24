using System;
using System.IO;

namespace ExcelCombiner
{
    static class Logger
    {
        public const string _EXTENSION = ".log";

        static StreamWriter m_writer;
        static StreamWriter m_standardWriter;

        public static void Initialize()
        {
            Initialize(DateTime.Now.ToString("yyyyMMdd"));
        }

        public static void Initialize(string filename)
        {
            m_writer = new StreamWriter(filename + _EXTENSION, true, Defines._SHIFT_JS);
            m_standardWriter = new StreamWriter(Console.OpenStandardOutput(), Defines._SHIFT_JS);
        }

        public static void Shutdown()
        {
            if (m_writer != null)
            {
                m_writer.Dispose();
                m_writer = null;
            }

            if (m_standardWriter != null)
            {
                m_standardWriter.Dispose();
                m_standardWriter = null;
            }
        }

        public static void WriteLine(string text, params object[] args)
        {
            Console.SetOut(m_writer);
            Console.WriteLine(text, args);
            m_writer.Flush();

            //Console画面へ出力 
            Console.SetOut(m_standardWriter);
            Console.WriteLine(text, args);
            m_standardWriter.Flush();
        }
    }
}
