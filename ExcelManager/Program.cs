using System;
using System.IO;
using System.Threading.Tasks;
using System.IO.Pipes;
using LinqToExcel;
using System.Collections.Generic;
using System.Reflection;

namespace ExcelManager
{
    class Program
    {

        /*
         *  C:\Users\FYBEDI\source\repos\ExcelManager\ExcelManager\bin\Debug\netcoreapp3.1\ExcelManager.exe C:\Users\FYBEDI\Desktop\lista.xlsx ListaGier B1 I5 wyszukajWiersz Premiera=2015,Platformy=PC
         * 
         */

        static public bool isFileFormatCorrect(string[] args)
        {
            if (args.Length != 6)
            {
                Console.WriteLine("Zła konstrukcja zapytania. Odpowiednia: '/Z1.exe <plik_WE.txt> <plik_WY.txt>");
                return false;
            }
            else if (!File.Exists(args[0]))
            {
                Console.WriteLine("Zła ścieżka do pliku wejściowego");
                return false;
            }
            else if (args[0] == args[1])
            {
                Console.WriteLine("Plik wejściowy i wyjściowy nie może być taki sam");
                return false;
            }
            return true;
        }

        static void Main(string[] args)
        {
            if (isFileFormatCorrect(args))
            {
                ExcelManagerTester excelManagerTester = new ExcelManagerTester(args);
                excelManagerTester.testAllFunctions();

            }
            
            /*
            using (NamedPipeServerStream pipeServer = new NamedPipeServerStream("PipeExcelConnect"))
            {
                Console.WriteLine("Utworzono serwer. Oczekiwanie na połączenie...");
                pipeServer.WaitForConnection();
                Console.WriteLine("Połączono!");
                StreamReader reader = new StreamReader(pipeServer);
                StreamWriter writer = new StreamWriter(pipeServer);
            }*/
        }
    }
}
