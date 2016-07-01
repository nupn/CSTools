using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.OleDb;


namespace ExcelByOLEDB
{
    class Program
    {
        static void Main(string[] args)
        {
            // 엑셀 문서 내용 추출
            String strFilePath = "../../test.xlsx";
            object missing = System.Reflection.Missing.Value;
            String valueString = "";

            string strProvider = string.Empty;
            strProvider = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" + strFilePath + @";Extended Properties=Excel 12.0";

            OleDbConnection excelConnection = new OleDbConnection(strProvider);
            excelConnection.Open();

            string strQuery = "SELECT * FROM [Sheet1$] where value > 1";

            OleDbCommand dbCommand = new OleDbCommand(strQuery, excelConnection);
            OleDbDataAdapter dataAdapter = new OleDbDataAdapter(dbCommand);

            DataTable dTable = new DataTable();
            dataAdapter.Fill(dTable);

            // dTable에 추출된 내용을 String으로 변환
            foreach (DataRow row in dTable.Rows)
            {
                foreach (DataColumn Col in dTable.Columns)
                {

                    Console.WriteLine(row[Col].ToString());

                    valueString += row[Col].ToString() + " ";
                }
            }

            dTable.Dispose();
            dataAdapter.Dispose();
            dbCommand.Dispose();

            excelConnection.Close();
            excelConnection.Dispose();

            Console.WriteLine(valueString);
            Console.ReadKey();

        }
    }
}
