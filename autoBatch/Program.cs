using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop;
using System.Threading.Tasks;
using System.Data.SqlClient;

namespace autoBatch
{
    class Program
    {
        static void Main(string[] args)
        {
            string path = @"\\DESIGNSVR1\dropbox\FINNPRODUCTION_temp.csv";
            Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(path);
            Microsoft.Office.Interop.Excel.Worksheet xlWorksheet = xlWorkbook.Sheets[1]; // assume it is the first sheet
            Microsoft.Office.Interop.Excel.Range xlRange = xlWorksheet.UsedRange; // get the entire used range
            int value = 10; //8 columns i guess --maybe 7
            int numberOfColumnsToRead = value;
            int last = xlWorksheet.Cells.SpecialCells(Microsoft.Office.Interop.Excel.XlCellType.xlCellTypeLastCell, Type.Missing).Row;
            Microsoft.Office.Interop.Excel.Range range = xlWorksheet.get_Range("A1:A" + last);


            //single connection here
            using (SqlConnection conn = new SqlConnection(CONNECT.ConnectionString))
            {
                conn.Open();
                //delete entried in that table here?
                string sql = "DELETE FROM dbo.auto_batch_finn_csv_import ";
                using (SqlCommand cmd = new SqlCommand(sql, conn))
                    cmd.ExecuteNonQuery();
                    for (int i = 1; i < last + 1; i++)
                    {
                        double temp = 0;
                        if (xlRange.Cells[i, 1].Value2 != null)
                        {
                            string charCheck = Convert.ToString(xlRange.Cells[i, 1].Value2);
                            if (charCheck.All(char.IsDigit))// if (System.Text.RegularExpressions.Regex.IsMatch(charCheck, @"^[a-zA-Z]+$") == false)
                                temp = xlRange.Cells[i, 1].Value2;
                        }
                        sql = "INSERT INTO dbo.auto_batch_finn_csv_import (door_id,program_id,material,thickness,length,width,quantity,date,machine) VALUES (" +
                   temp + ",'" + //door id
                   xlRange.Cells[i, 2].Value2.ToString() + "','" +//program id
                   xlRange.Cells[i, 3].Value2.ToString() + "'," +//material
                   xlRange.Cells[i, 4].Value2.ToString() + "," +//thickness
                   xlRange.Cells[i, 5].Value2.ToString() + "," +//length
                   xlRange.Cells[i, 6].Value2.ToString() + "," +//width
                   xlRange.Cells[i, 7].Value2.ToString() + ",'" +//quantity
                   xlRange.Cells[i, 8].Value2.ToString() + "','" +//date
                   xlRange.Cells[i, 9].Value2.ToString() + "')";//machine
                        using (SqlCommand cmd = new SqlCommand(sql, conn))
                            cmd.ExecuteNonQuery();
                        Console.WriteLine("row: " + i.ToString() + " inserted :}");

                        //if (xlRange.Cells[i, 1].Value2 != null)
                        //    Console.WriteLine(xlRange.Cells[i, 1].Value2.ToString()); // do whatever with value
                    }
                conn.Close();
            }


            xlWorkbook.Close(0);
            xlApp.Quit();

        }
    }
}
