using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelCombiner
{
    static class Defines
    {
        public const int _INVALID_VALUE = -1;

        public const string _EXTENSION_XLS = ".xls";
        public const string _EXTENSION_XLSX = ".xlsx";
        public const string _EXTENSION_CS = ".cs";
        public const string _OUTPUT_FILE_NAME = "ScenarioTable.xls";

        public static readonly System.Text.Encoding _SHIFT_JS = Encoding.GetEncoding("Shift_JIS");
    }
}
