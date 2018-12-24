SET CONVERTER=%ExcelCombiner_Root%\Executables\ExcelCombiner.exe
SET DATA_DIR=%ExcelCombiner_Root%\Data\\
SET FILENAME=ScenarioTable.xls
SET OUTPUT_DIR=%ExcelCombiner_Root%\%FILENAME%

"%CONVERTER%" "%DATA_DIR%" "%OUTPUT_DIR%"
