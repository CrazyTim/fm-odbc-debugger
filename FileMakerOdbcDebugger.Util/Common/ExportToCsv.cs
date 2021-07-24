using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Text;

namespace FileMakerOdbcDebugger.Util
{
    public static partial class Common
    {
        private static long fileNameIncrementor = 0;
        private static readonly string tempFolder = Path.GetTempPath() + "fm-odbc-debugger";

        public static void ExportToCsv(List<List<string>> data)
        {
            var pathToFile = GetPathToFile();
            
            using (var stream = File.OpenWrite(pathToFile))
            {
                using (var streamWriter = new StreamWriter(stream, Encoding.UTF8))
                {
                    using (var csvWriter = new CsvHelper.CsvWriter(streamWriter, CultureInfo.InvariantCulture))
                    {
                        foreach (var row in data)
                        {
                            foreach (var i in row)
                            {
                                csvWriter.WriteField(i);
                            }
                            csvWriter.NextRecord();
                        }
                    }
                }
            }

            Process.Start(pathToFile);
        }

        private static string GetPathToFile()
        {
            Directory.CreateDirectory(tempFolder);

            string pathToFile;
            do
            {
                pathToFile = $@"{tempFolder}\export-{fileNameIncrementor}.csv";
                fileNameIncrementor += 1;
            }
            while (File.Exists(pathToFile));

            return pathToFile;
        }
    }
}
