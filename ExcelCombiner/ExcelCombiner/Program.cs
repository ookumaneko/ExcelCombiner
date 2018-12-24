using System;
using System.IO;
using System.Linq;

namespace ExcelCombiner
{
    class Program
    {
        enum PathIndex
        {
            Load = 0,
            Output,
            Count
        }

        static void Main(string[] args)
        {
            Logger.Initialize();

            string directory = string.Empty;
            string outputDirectory = string.Empty;
            int argLength = args.Length;

            if (argLength < (int)PathIndex.Count)
            {
                // パスの数が足りない場合は、取りあえず今のパスを変わりに使う
                directory = Directory.GetCurrentDirectory();
                outputDirectory = Directory.GetCurrentDirectory() + "/" + Defines._OUTPUT_FILE_NAME;
                Logger.WriteLine("Not enough path defined : count = " + argLength);
            }
            else
            {
                // 順番間違ってたら死
                directory = args[(int)PathIndex.Load];
                outputDirectory = args[(int)PathIndex.Output];
            }

            // 取りあえずパスのログ出す
            for (int i = 0; i < argLength; ++i)
            {
                Logger.WriteLine(string.Format("arg[{0}]:{1}", i, args[i]));
            }

            Converter converter = new Converter(directory, outputDirectory);
            var files = Directory.GetFiles(directory, "*.*", SearchOption.AllDirectories).Where(s => s.EndsWith(Defines._EXTENSION_XLS));
            converter.Convert(files.ToArray());

            Logger.WriteLine("{0}{1}{2}", Environment.NewLine, "------ Finished Converting!", Environment.NewLine );
            Logger.Shutdown();
            Console.ReadLine();            
        }
    }
}
