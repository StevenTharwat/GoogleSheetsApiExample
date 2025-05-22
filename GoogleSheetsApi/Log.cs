using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GoogleSheetsApi
{
    public static class Log
    {
        static Log()
        {
            write("\n -------------------------------------------------- Start a new App -------------------------------------------------- \n");
        }

        public static bool write(string message)
        {
            try
            {
                string path = "E:\\projects\\libraries\\GoogleSheetsApi\\GoogleSheetsApi\\GoogleSheetsApi\\log.txt";
                using (StreamWriter writer = new StreamWriter(path, true))
                {
                    writer.WriteLine($"{DateTime.Now} : {message}");
                }
                return true;
            }
            catch
            {
                return false;
            }
        }
    }
}
