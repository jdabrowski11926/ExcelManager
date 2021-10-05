using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using LinqToExcel;
using System.Linq;

namespace ExcelManager
{

    class ExcelManager
    {

        /*************************************************************************************/
        /* PLIK - operacje na pliku
         /*************************************************************************************/

        public ExcelQueryFactory getExcelFile(string filePath)
        {
            if (!File.Exists(filePath)) throw new Exception("Nie znaleziono pliku o podanej ścieżce");
            return new ExcelQueryFactory(filePath);
        }

        /*************************************************************************************/
        /* WORKSHEET - operacje na arkuszach
         /*************************************************************************************/

        public List<string> getExcelWorksheetNames(string filePath)
        {
            ExcelQueryFactory excelFile = getExcelFile(filePath);
            return excelFile.GetWorksheetNames().ToList();
        }

        public string getExcelWorksheetNameAtIndex(string filePath, int index)
        {
            if (index < 0 || index >= getExcelWorksheetNames(filePath).Count()) throw new Exception("Nie ma arkusza o podanym indeksie");
            return getExcelWorksheetNames(filePath)[index];
        }

        public void printExcelWorksheetNames(string filePath)
        {
            List<string> worksheetNames = getExcelWorksheetNames(filePath);
            printList<string>(worksheetNames, "Nazwy arkuszy:");
        }

        public bool doesWorksheetExsist(string filePath, string worksheetName)
        {
            if (!getExcelWorksheetNames(filePath).Contains(worksheetName)) 
                return false;
            return true;
        }

        /*************************************************************************************/
        /* COLUMN - operacje na kolumnach
         /*************************************************************************************/

        public List<string> getExcelWorksheetColumnNames(string filePath, string worksheetName)
        {
            ExcelQueryFactory excelFile = getExcelFile(filePath);
            return excelFile.GetColumnNames(worksheetName).ToList();
        }

        public List<string> getExcelWorksheetColumnNames(string filePath)
        {
            return getExcelWorksheetColumnNames(filePath, getExcelWorksheetNameAtIndex(filePath, 0));
        }

        public void printExcelWorksheetColumnNames(string filePath, string worksheetName)
        {
            List<string> columnNames = getExcelWorksheetColumnNames(filePath, worksheetName);
            printList<string>(columnNames, "Nazwy kolumn:");
        }

        public void printExcelWorksheetColumnNames(string filePath)
        {
            printExcelWorksheetColumnNames(filePath, getExcelWorksheetNameAtIndex(filePath, 0));
        }

        /*************************************************************************************/
        /* DATA - pobieranie danych z arkusza
         /*************************************************************************************/

        public List<Row> getExcelWorksheetData(string filePath)
        {
            return getExcelWorksheetData(filePath, getExcelWorksheetNameAtIndex(filePath, 0), new List<string>());
        }

        public List<Row> getExcelWorksheetData(string filePath, string worksheetName)
        {
            return getExcelWorksheetData(filePath, worksheetName, new List<string>());
        }

        public List<Row> getExcelWorksheetData(string filePath, string worksheetName, List<string> requirements)
        {
            ExcelQueryFactory excelFile = getExcelFile(filePath);
            if(!doesWorksheetExsist(filePath, worksheetName))
                throw new Exception("Nie znaleziono arkusza o podanej nazwie");

            var worksheetData = from c in excelFile.Worksheet(worksheetName) select c;
            if (requirements?.Any()==false) 
                return worksheetData.ToList();
            return getFilteredRows(worksheetData.ToList(), requirements, getExcelWorksheetColumnNames(filePath, worksheetName));
        }

        public List<Row> getExcelWorksheetData(string filePath, string rangeStart, string rangeEnd)
        {
            return getExcelWorksheetData(filePath, getExcelWorksheetNameAtIndex(filePath, 0), rangeStart, rangeEnd, new List<string>());
        }

        public List<Row> getExcelWorksheetData(string filePath, string worksheetName, string rangeStart, string rangeEnd)
        {
            return getExcelWorksheetData(filePath, worksheetName, rangeStart, rangeEnd, new List<string>());
        }

        public List<Row> getExcelWorksheetData(string filePath, string worksheetName, string rangeStart, string rangeEnd, List<string> requirements)
        {
            ExcelQueryFactory excelFile = getExcelFile(filePath);
            if (!doesWorksheetExsist(filePath, worksheetName))
                throw new Exception("Nie znaleziono arkusza o podanej nazwie");

            var worksheetData = from c in excelFile.WorksheetRange(rangeStart, rangeEnd, worksheetName) select c;
            if (requirements?.Any() == false)
                return worksheetData.ToList();
            return getFilteredRows(worksheetData.ToList(), requirements, getExcelWorksheetColumnNames(filePath, worksheetName));
        }

        public void printExcelWorksheetData(string filePath)
        {
            printRows(getExcelWorksheetData(filePath));
        }

        public void printExcelWorksheetData(string filePath, string worksheetName)
        {
            printRows(getExcelWorksheetData(filePath, worksheetName));
        }

        public void printExcelWorksheetData(string filePath, string worksheetName, List<string> requirements)
        {
            printRows(getExcelWorksheetData(filePath, worksheetName, requirements));
        }

        public void printExcelWorksheetData(string filePath, string rangeStart, string rangeEnd)
        {
            printRows(getExcelWorksheetData(filePath, rangeStart, rangeEnd));
        }

        public void printExcelWorksheetData(string filePath, string worksheetName, string rangeStart, string rangeEnd)
        {
            printRows(getExcelWorksheetData(filePath, worksheetName, rangeStart, rangeEnd));
        }

        public void printExcelWorksheetData(string filePath, string worksheetName, string rangeStart, string rangeEnd, List<string> requirements)
        {
            printRows(getExcelWorksheetData(filePath, worksheetName, rangeStart, rangeEnd, requirements));
        }

        /*************************************************************************************/
        /* REQUIREMENS - analizowanie wymagań
         /*************************************************************************************/

        public List<Row> getFilteredRows(List<Row> entryRows, List<string> requirements, List<string> columnNames)
        {
            List<Row> allowedRows = new List<Row>();
            foreach (Row row in entryRows)
            {
                if (isRowAllowed(row, requirements, columnNames))
                    allowedRows.Add(row);
            }
            return allowedRows;
        }

        public bool isRowAllowed(Row row, List<string> requirements, List<string> columnNames)
        {
            for (int i = 0; i < requirements.Count; i++)
            {
                if (!isRequirementFulfilled(row, requirements[i], columnNames))
                    return false;
            }
            return true;
        }

        public bool isRequirementFulfilled(Row row, string requirement, List<string> columnNames)
        {
            var requirementSplit = requirement.Split('=');
            string columnName = requirementSplit[0];
            string requiredValue = requirementSplit[1];
            int columnIndex = getColumnIndex(columnName, columnNames);
            if (!row[columnIndex].ToString().Contains(requiredValue))
                return false;
            return true;
        }

        public int getColumnIndex(string columnName, List<string> columnNames)
        {
            for (int i = 0; i < columnNames.Count; i++)
                if (columnName.Equals(columnNames[i])) return i;
            return -1;
        }

        /*************************************************************************************/
        /* FUNKCJE PRINT ROWS AND CELLS - Wyświetlanie wierszy oraz komórek z excella (funkcje pomocnicze) */
        /*************************************************************************************/

        public void printList<T>(List<T> list, string title = null)
        {
            if (title != null) Console.WriteLine(title);
            for (int i = 0; i < list.Count; i++)
                Console.WriteLine(list[i]);
        }

        public void printRows(List<RowNoHeader> worksheetData)
        {
            foreach (RowNoHeader row in worksheetData)
                printRow(row);
        }

        public void printRows(List<Row> worksheetData)
        {
            foreach (Row row in worksheetData)
                printRow(row);
        }

        public void printRow(RowNoHeader row)
        {
            foreach (Cell cell in row)
                printCell(cell);
            Console.WriteLine("");
        }

        public void printRow(Row row)
        {
            foreach (Cell cell in row)
                printCell(cell);
            Console.WriteLine("");
        }

        public void printCell(Cell cell)
        {
            Console.Write(cell.ToString() + "   ||   ");
        }

        /*************************************************************************************/

    }
}
