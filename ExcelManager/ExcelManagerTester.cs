using LinqToExcel;
using System;
using System.Collections.Generic;
using System.Reflection;
using System.Text;
using System.Linq;

namespace ExcelManager
{
    class ExcelManagerTester
    {
        static String filePath, worksheetName, rangeStart, rangeEnd, commandType;
        ExcelManager excelManager;
        List<string> requirements;

        public ExcelManagerTester(string[] args)
        {
            filePath = args[0];
            worksheetName = args[1];
            rangeStart = args[2];
            rangeEnd = args[3];
            commandType = args[4];
            requirements = args[5].Split(',').ToList();
            excelManager = new ExcelManager();
        }

        public void testAllFunctions()
        {
            testGetWorksheetNames();
            testGetWorksheetColumnNames();
            testWorksheetGetData();
        }

        public void executeFunction()
        {
            switch (commandType)
            {
                case "getWorksheetNames":

                    break;

                case "getColumnNames":

                    break;

                case "getData":

                    break;
            }
        }
        
        public void printFunctionTest(string function, object[] parameters)
        {
            Console.WriteLine("######################################################################################################");
            Console.Write("Testowanie funkcji '" + function + "' z parametrami ");
            printParameters(parameters);
            testFunction(function, parameters);
            Console.WriteLine("######################################################################################################");
        }

        public void testFunction(string function, object[] parameters)
        {
            try
            {
                MethodInfo mi = excelManager.GetType().GetMethod(function, getMethodTypes(parameters));
                mi.Invoke(excelManager, parameters);
            }
            catch (Exception e)
            {
                Console.WriteLine("WYSTĄPIŁ BŁĄD : " + e.Message);
            }
        }

        public void printParameters(object[] parameters)
        {
            Console.Write("(");
            for (int i = 0; i < parameters.Length; i++)
            {
                Console.Write(parameters[i].ToString());
                if (i != parameters.Length - 1 || parameters.Length == 0)
                    Console.Write(", ");
            }
            Console.WriteLine(")");
        }

        public Type[] getMethodTypes(object[] parameters)
        {
            List<Type> typeList = new List<Type>();
            foreach (object parameter in parameters)
                typeList.Add(parameter.GetType());  
            return typeList.ToArray();
        }

        public void testGetWorksheetNames()
        {
            object[] parameters = new object[] { filePath };
            printFunctionTest("printExcelWorksheetNames", parameters);

            parameters = new object[] { "wrongFilePath" };
            printFunctionTest("printExcelWorksheetNames", parameters);
        }

        public void testGetWorksheetColumnNames()
        {
            object[] parameters = new object[] { filePath };
            printFunctionTest("printExcelWorksheetColumnNames", parameters);

            parameters = new object[] { filePath, worksheetName };
            printFunctionTest("printExcelWorksheetColumnNames", parameters);

            parameters = new object[] { filePath, "wrongWorksheetName" };
            printFunctionTest("printExcelWorksheetColumnNames", parameters);
        }

        public void testWorksheetGetData()
        {
            object[] parameters = new object[] { filePath };
            printFunctionTest("printExcelWorksheetData", parameters);

            parameters = new object[] { filePath, worksheetName };
            printFunctionTest("printExcelWorksheetData", parameters);

            parameters = new object[] { filePath, worksheetName, requirements };
            printFunctionTest("printExcelWorksheetData", parameters);

            parameters = new object[] { filePath, rangeStart, rangeEnd };
            printFunctionTest("printExcelWorksheetData", parameters);

            parameters = new object[] { filePath, worksheetName, rangeStart, rangeEnd };
            printFunctionTest("printExcelWorksheetData", parameters);

            parameters = new object[] { filePath, worksheetName, rangeStart, rangeEnd, requirements };
            printFunctionTest("printExcelWorksheetData", parameters);
        }

    }
}
