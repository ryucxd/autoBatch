using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop;
using System.Threading.Tasks;
using System.Data.SqlClient;
using System.Data;

namespace autoBatch
{
    class Program
    {
        public int _current_number { get; set; }
        public int _limit { get; set; }
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
            string sql = "";

            //single connection here
            using (SqlConnection conn = new SqlConnection(CONNECT.ConnectionString))
            {
                conn.Open();
                //delete entried in that table here?
                sql = "DELETE FROM dbo.auto_batch_finn_csv_import ";
                using (SqlCommand cmd = new SqlCommand(sql, conn))
                    cmd.ExecuteNonQuery();
                for (int i = 1; i < last; i++)
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


            //at this point now we would run the sql procedure
            sql = "SELECT [current_batch_no],[limit] FROM dbo.auto_batch_limit";
            using (SqlConnection conn = new SqlConnection(CONNECT.ConnectionString))
            {
                conn.Open();
                int current_number = 0;
                int limit = 0;
                using (SqlCommand cmd = new SqlCommand(sql, conn)) //first up confirm that the current batch is < the limiter
                {
                    SqlDataAdapter da = new SqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    da.Fill(dt);
                    foreach (DataRow row in dt.Rows)
                    {
                        current_number = Convert.ToInt32(row["current_batch_no"].ToString());
                        limit = Convert.ToInt32(row["limit"].ToString());
                    }
                }
                if (current_number < limit)
                {
                    while (current_number < limit)
                    {
                        //run auto_batch_master  and then check for new limit x current
                        using (SqlCommand cmdUSP = new SqlCommand("auto_batch_master", conn))
                        {
                            cmdUSP.CommandType = CommandType.StoredProcedure;
                            cmdUSP.ExecuteNonQuery();
                        }



                        sql = "SELECT [current_batch_no],[limit] FROM dbo.auto_batch_limit";
                        using (SqlCommand cmd = new SqlCommand(sql, conn)) //lastly confirm that the current batch is still < the limiter
                        {
                            SqlDataAdapter da = new SqlDataAdapter(cmd);
                            DataTable dt = new DataTable();
                            da.Fill(dt);
                            foreach (DataRow row in dt.Rows)
                            {
                                current_number = Convert.ToInt32(row["current_batch_no"].ToString());
                                limit = Convert.ToInt32(row["limit"].ToString());
                                Console.WriteLine("----");
                                Console.WriteLine("Current batch list total number: " + current_number.ToString());
                                Console.WriteLine("----");
                            }
                        }
                        sql = "UPDATE dbo.auto_batch_limit SET current_batch_no = current_batch_no  + 1 ";
                        using (SqlCommand cmdCurrentBatchNo = new SqlCommand(sql, conn))
                            cmdCurrentBatchNo.ExecuteNonQuery();
                    }
                }
                else return;
                conn.Close();
                startBatch();
                Console.WriteLine("End of batch...");
                Console.ReadLine();
            }
        }


        private static void startBatch()
        {
            //at this point we can assume that the  auto_batch master has dumped some info into [auto_batch_selected_door]
            //and if it hasnt then end the loop entirely 
            string sql = "SELECT COUNT(id) FROM dbo.auto_batch_selected_door"; //fill a datatable   
            int count = 0;
            using (SqlConnection connCount = new SqlConnection(CONNECT.ConnectionString))
            {
                connCount.Open();
                using (SqlCommand cmdCheckData = new SqlCommand(sql, connCount))
                {
                    count = Convert.ToInt32(cmdCheckData.ExecuteScalar());
                    if (count < 0)
                        return;
                }
                connCount.Close();
            }

            //this is where things start getting giga confusing....
            //ok so i think the first step is inserting into dbo.batch


            DataTable dt = new DataTable();
            using (SqlConnection conn = new SqlConnection(CONNECT.ConnectionString))
            {
                while (count > 0)
                {
                    conn.Open();


                    sql = "SELECT * from dbo.auto_batch_selected_door";
                    using (SqlCommand cmdDT = new SqlCommand(sql, conn))
                    {
                        SqlDataAdapter da = new SqlDataAdapter(cmdDT);
                        da.Fill(dt);
                    }
                    //couple of variables we need for the insert etc
                    int door_id = 0, batch_id = 0, finn = 0, rainer = 0, quantity = 0;
                    string program_no = "";  //looks like program no is only used when part batching? same with qty
                    door_id = Convert.ToInt32(dt.Rows[0][1].ToString());
                    program_no = dt.Rows[0][2].ToString();
                    quantity = Convert.ToInt32(dt.Rows[0][7].ToString());
                    if (dt.Rows[0][9].ToString() == "RAINER")
                    {
                        rainer = -1;
                        finn = 0;
                    }
                    else if (dt.Rows[0][9].ToString() == "FINN-POWER")
                    {
                        rainer = 0;
                        finn = -1;
                    }

                    //get max batch id rq
                    using (SqlCommand cmdBatchID = new SqlCommand("Select MAX(batch_id) + 1 FROM dbo.batch", conn))
                        batch_id = Convert.ToInt32(cmdBatchID.ExecuteScalar());
                    //set batch id to something static because we want to replicate a finished search
                    batch_id = 9692;

                    sql = "INSERT INTO dbo.batch (door_id,batch_date,batch_id,part_batched,offcut) VALUES" +
                        "(" + door_id.ToString() + ",GETDATE()," + batch_id.ToString() + ",0,0)";
                    Console.WriteLine(sql);

                    //next step in the batch is to set finn OR rainer in dbo.door
                    //Case "FINN-POWER"
                    if (finn == -1)
                        sql = "UPDATE dbo_door set finn = finn + " + quantity + " where id=" + +door_id;
                    else if (rainer == -1)
                        sql = "UPDATE dbo_door set rainer = rainer + " + quantity + " where id= " + door_id;
                    else
                        sql = "error no machine -1";
                    Console.WriteLine(sql);

                    //set as batched
                    sql = "update dbo.door set batch_live =0, batched=-1 where id =" + door_id;
                    Console.WriteLine(sql);
                    count = count - 1;
                    conn.Close();
                }
                WriteToPunchProgQ();
            }
        }

        private static void WriteToPunchProgQ()
        {
            //check if there is anything for the finn(?)
            string sql = "";
            int temp = 0;
            int batch_id = 0;
            string grouping = "";
            DataTable finnBatchDT = new DataTable();
            using (SqlConnection conn = new SqlConnection(CONNECT.ConnectionString))
            {
                conn.Open();
                sql = "Select * FROM dbo.auto_batch_selected_door WHERE machine = 'FINN-POWER'";
                DataTable dt = new DataTable();
                using (SqlCommand cmd = new SqlCommand(sql, conn))
                {
                    SqlDataAdapter da = new SqlDataAdapter(cmd);
                    da.Fill(dt);
                    temp = Convert.ToInt32(dt.Rows[0][0].ToString());
                }
                if (temp == 0) //need to test this 
                    return; //there is nothing for the finn

                //insert into bath header  (grab the max batch_id + 1 )
                sql = "SELECT MAX(batch_id) + 1 FROM dbo.batch";
                using (SqlCommand cmd = new SqlCommand(sql, conn))
                    batch_id = Convert.ToInt32(cmd.ExecuteNonQuery());

                //set batch id to something static because we want to replicate a finished search
                batch_id = 9692;

                //loop here for EACH program number of EACH door that needs to be batched
                int doorCount = 0;
                int counter = 0;
                sql = "select count(id) from dbo.auto_batch_selected_door ";
                using (SqlCommand cmd = new SqlCommand(sql, conn))
                    doorCount = Convert.ToInt32(cmd.ExecuteScalar());
                while (counter < doorCount)
                {
                    //grouping
                    sql = "SELECT [group ]FROM dbo.auto_batch_finn_batch where program_id = '" + dt.Rows[counter][2] + "'";
                    using (SqlCommand cmd = new SqlCommand(sql, conn))
                        grouping = Convert.ToString(cmd.ExecuteScalar());

                    //batchheader
                    sql = "insert into dbo_batch_header (qid, qname, datecreated, machine) values ('" + batch_id + "','" + dt.Rows[counter][9] + "','" + DateTime.Now.ToString() + "','Finn-Power');";

                    Console.WriteLine(sql);
                    Console.WriteLine("--"); ;

                    // i think this can largely be ignored by just reading the datatable from above... i think?
                    sql = "SELECT dbo.batch_header.id, dbo.batch_header.qid, dbo.batch_header.qname, auto_batch_finn_batch.program_id, auto_batch_finn_batch.FirstOfquantity, auto_batch_finn_batch.door_id " +
                        "FROM auto_batch_finn_batch " +
                        "INNER JOIN dbo.batch_header ON auto_batch_finn_batch.[group] = dbo.batch_header.qname " +
                        "WHERE [group] = '" + grouping + "' " +
                        "GROUP BY dbo.batch_header.id, dbo.batch_header.qid, dbo.batch_header.qname, auto_batch_finn_batch.program_id, auto_batch_finn_batch.FirstOfquantity, auto_batch_finn_batch.door_id";
                    using (SqlCommand cmd = new SqlCommand(sql, conn))
                    {
                        SqlDataAdapter da = new SqlDataAdapter(cmd);
                        da.Fill(finnBatchDT);
                    }

                        Console.WriteLine(sql);
                    Console.WriteLine("--"); ;

                    sql = "insert into dbo_batch_programs (door_id, program_no, sheet_quantity, header_id) values ('" + finnBatchDT.Rows[counter][5].ToString() + "','" + finnBatchDT.Rows[counter][3].ToString() + "','" + finnBatchDT.Rows[counter][4].ToString() + "'," + finnBatchDT.Rows[counter][0].ToString() + ");";

                          Console.WriteLine(sql);
                    Console.WriteLine("--"); ;


                    counter = counter + 1;
                    Console.WriteLine("End of WriteToPunchProgQ loop - press enter to continue");
                    Console.ReadLine(); //pause
                }
                conn.Close();
            }
        }

        private static void writeToPunchProgQRainer()
        {

            //check if there is anything for the finn(?)
            string sql = "";
            int temp = 0;
            int batch_id = 0;
            string grouping = "";
            DataTable finnBatchDT = new DataTable();
            using (SqlConnection conn = new SqlConnection(CONNECT.ConnectionString))
            {
                conn.Open();
                sql = "Select * FROM dbo.auto_batch_selected_door WHERE machine = 'RAINER'";
                DataTable dt = new DataTable();
                using (SqlCommand cmd = new SqlCommand(sql, conn))
                {
                    SqlDataAdapter da = new SqlDataAdapter(cmd);
                    da.Fill(dt);
                    temp = Convert.ToInt32(dt.Rows[0][0].ToString());
                }
                if (temp == 0) //need to test this 
                    return; //there is nothing for the finn

                //insert into bath header  (grab the max batch_id + 1 )
                sql = "SELECT MAX(batch_id) + 1 FROM dbo.batch";
                using (SqlCommand cmd = new SqlCommand(sql, conn))
                    batch_id = Convert.ToInt32(cmd.ExecuteNonQuery());

                //set batch id to something static because we want to replicate a finished search
                batch_id = 9692;

                //loop here for EACH program number of EACH door that needs to be batched
                int doorCount = 0;
                int counter = 0;
                sql = "select count(id) from dbo.auto_batch_selected_door ";
                using (SqlCommand cmd = new SqlCommand(sql, conn))
                    doorCount = Convert.ToInt32(cmd.ExecuteScalar());
                while (counter < doorCount)
                {
                    //grouping
                    sql = "SELECT [group ]FROM dbo.auto_batch_rainer_batch where program_id = '" + dt.Rows[counter][2] + "'";
                    using (SqlCommand cmd = new SqlCommand(sql, conn))
                        grouping = Convert.ToString(cmd.ExecuteScalar());

                    //batchheader
                    sql = "insert into dbo_batch_header (qid, qname, datecreated, machine) values ('" + batch_id + "','" + dt.Rows[counter][9] + "','" + DateTime.Now.ToString() + "','Finn-Power');";

                    Console.WriteLine(sql);
                    Console.WriteLine("--"); ;

                    // i think this can largely be ignored by just reading the datatable from above... i think?
                    sql = "SELECT dbo.batch_header.id, dbo.batch_header.qid, dbo.batch_header.qname, auto_batch_rainer_batch.program_id, auto_batch_rainer_batch.FirstOfquantity, auto_batch_rainer_batch.door_id " +
                        "FROM auto_batch_rainer_batch " +
                        "INNER JOIN dbo.batch_header ON auto_batch_rainer_batch.[group] = dbo.batch_header.qname " +
                        "WHERE [group] = '" + grouping + "' " +
                        "GROUP BY dbo.batch_header.id, dbo.batch_header.qid, dbo.batch_header.qname, auto_batch_rainer_batch.program_id, auto_batch_rainer_batch.FirstOfquantity, auto_batch_rainer_batch.door_id";
                    using (SqlCommand cmd = new SqlCommand(sql, conn))
                    {
                        SqlDataAdapter da = new SqlDataAdapter(cmd);
                        da.Fill(finnBatchDT);
                    }

                    Console.WriteLine(sql);
                    Console.WriteLine("--"); ;

                    sql = "insert into dbo_batch_programs (door_id, program_no, sheet_quantity, header_id) values ('" + finnBatchDT.Rows[counter][5].ToString() + "','" + finnBatchDT.Rows[counter][3].ToString() + "','" + finnBatchDT.Rows[counter][4].ToString() + "'," + finnBatchDT.Rows[counter][0].ToString() + ");";

                    Console.WriteLine(sql);
                    Console.WriteLine("--"); ;


                    counter = counter + 1;
                    Console.WriteLine("End of WriteToPunchProgQ loop - press enter to continue");
                    Console.ReadLine(); //pause
                }
                conn.Close();
            }
        }  //this cant really be tested because there are 0 entried of 'mild steel in the csv'


        private static void autoWriteToFinn()
        {
            string sql = "select batch_id from dbo.auto_batch_finn_batch where FirstOfmachine = 'FINN-POWER' group by batch_id ";
            DataTable dtBatchIdList = new DataTable();
            using (SqlConnection conn = new SqlConnection(CONNECT.ConnectionString))
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand(sql, conn))
                {
                    SqlDataAdapter da = new SqlDataAdapter(cmd);
                    da.Fill(dtBatchIdList);
                }

                //now we loop through the DT

                conn.Close();
            }
        }

    }
}

