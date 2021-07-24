using System;

namespace FileMakerOdbcDebugger.Util
{
    public static partial class Common
    {
        /// <summary>
        /// Return the file version of the driver 'fmodbc64.dll' located in the System32 folder.
        /// Return an empty string if the file doesn't exist.
        /// </summary>
        public static string GetOdbcDriverVersion64Bit()
        {
            var systemDrive = Environment.GetEnvironmentVariable("windir", EnvironmentVariableTarget.Machine);
            var pathToOdbcDriver = systemDrive + @"\System32\fmodbc64.dll";

            if (System.IO.File.Exists(pathToOdbcDriver))
            {
                var myFileVersionInfo = System.Diagnostics.FileVersionInfo.GetVersionInfo(pathToOdbcDriver);
                return myFileVersionInfo.FileVersion;
            }

            return "";
        }
    }
}
