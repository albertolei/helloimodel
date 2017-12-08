using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;


namespace NImodel.Utils
{
    class FileIOUtil
    {
        private static string filepath = @"D:\imodel_info\imodel.txt";

        public static void writeToFile(string content)
        {
            StreamWriter sw = new StreamWriter(filepath, true);
            sw.WriteLine(content);
            sw.Flush();
            sw.Close();
        }
    }
}
