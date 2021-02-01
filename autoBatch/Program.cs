using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop;
using System.Threading.Tasks;
using System.Data.SqlClient;
using System.Data;
using System.IO;
using System.Data.OleDb;

namespace autoBatch
{
    class Program
    {
        public int _current_number { get; set; }
        public int _limit { get; set; }
        static void Main(string[] args)
        {
            int bufferHeight = Console.BufferHeight;
            int bufferWidth = Console.BufferWidth;


            //bufferHeight += 240;
            //Console.BufferHeight = bufferHeight;

            bufferWidth += 85;
            Console.BufferWidth = bufferWidth;





            //autoWriteToFinn();
            //return;
            string path = @"\\DESIGNSVR1\dropbox\FINNPRODUCTION_temp.csv";                //@"\\DESIGNSVR1\subcontracts\Express5.MAC\FINNPRODUCTION.csv";
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
                //////////sql = "DELETE FROM dbo.auto_batch_finn_csv_import "; 
                //////////using (SqlCommand cmd = new SqlCommand(sql, conn))
                //////////    cmd.ExecuteNonQuery();
                for (int i = 1; i < 100; i++)  //for (int i = 1; i < last; i++) 
                {
                    double temp = 0;
                    if (xlRange.Cells[i, 1].Value2 != null)
                    {
                        string charCheck = Convert.ToString(xlRange.Cells[i, 1].Value2);
                        if (charCheck.All(char.IsDigit))// if (System.Text.RegularExpressions.Regex.IsMatch(charCheck, @"^[a-zA-Z]+$") == false)
                            temp = xlRange.Cells[i, 1].Value2;
                    }
                    //remove the space from the start of material here 
                    string tempString = xlRange.Cells[i, 3].Value2.ToString();
                    tempString = tempString.Trim();

                    sql = "INSERT INTO dbo.auto_batch_finn_csv_import (door_id,program_id,material,thickness,length,width,quantity,date,machine) VALUES (" +
                   temp + ",'" + //door id
                   xlRange.Cells[i, 2].Value2.ToString() + "','" +//program id
                   tempString + "'," +//material (with absouletly no leading or trailing white spaces!!! :DDD  (this was causing the views to break 
                   xlRange.Cells[i, 4].Value2.ToString() + "," +//thickness
                   xlRange.Cells[i, 5].Value2.ToString() + "," +//length
                   xlRange.Cells[i, 6].Value2.ToString() + "," +//width
                   xlRange.Cells[i, 7].Value2.ToString() + ",'" +//quantity
                   xlRange.Cells[i, 8].Value2.ToString() + "','" +//date
                   xlRange.Cells[i, 9].Value2.ToString() + "')";//machine
                    ////////////using (SqlCommand cmd = new SqlCommand(sql, conn))
                    ////////////    cmd.ExecuteNonQuery();
                    Console.WriteLine("row: " + i.ToString() + " inserted :}");

                    //if (xlRange.Cells[i, 1].Value2 != null)
                    //    Console.WriteLine(xlRange.Cells[i, 1].Value2.ToString()); // do whatever with value
                }
                conn.Close();
            }


            xlWorkbook.Close(0);
            xlApp.Quit(); //this isnt enough to properly close the app... 

            //this will loop through every process and kill anything that is related to excel - this is probably fine as it'll be run somewhere where there is no user opening excel files
            System.Diagnostics.Process[] process = System.Diagnostics.Process.GetProcessesByName("Excel"); 
            foreach (System.Diagnostics.Process p in process)
            {
                if (!string.IsNullOrEmpty(p.ProcessName))
                {
                    try
                    {
                        p.Kill(); //
                    }
                    catch { }
                }
            }


            //at this point now we would run the sql procedure  
            sql = "SELECT [current_batch_no],[limit] FROM dbo.auto_batch_limit";
            using (SqlConnection conn = new SqlConnection(CONNECT.ConnectionString))
            {
                conn.Open();
                int current_number = 0;
                int limit = 0;
                using (SqlCommand cmd = new SqlCommand(sql, conn)) //first up confirm that the current batch is < the limiter  //the limit is going an extra 1 somewhere below this
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
                if (current_number < limit) //i think this section needs to add one onto the current number
                {
                    while (current_number < limit)
                    {
                        //run auto_batch_master  and then check for new limit x current
                        using (SqlCommand cmdUSP = new SqlCommand("auto_batch_master", conn))
                        {
                            cmdUSP.CommandType = CommandType.StoredProcedure;
                            cmdUSP.ExecuteNonQuery();
                        }
                        //here we need to check to see if the procedure has found any doors > none are found then stop running the procedure

                        sql = "SELECT no_doors FROM dbo.auto_batch_limit";
                        using (SqlCommand cmd = new SqlCommand(sql, conn))
                        {
                            int doorCount = Convert.ToInt32(cmd.ExecuteScalar());
                            if (doorCount == -1) //there are NO more doors to batch so exit out
                            {
                                sql = "update dbo.auto_batch_limit SET no_doors = 0"; //set this back to 0 so that the next time it runs it doesnt hit it
                                using (SqlCommand cmd2 = new SqlCommand(sql, conn))
                                    cmd2.ExecuteScalar();
                                current_number = 99999999;
                                continue;
                            }
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
                else
                {
                    Console.WriteLine("Already max number of doors batching... press any key to exit "); //will need to remove all of these by the time we publish it! :}
                    Console.ReadLine();
                    Environment.Exit(-1); //exit out of the app~
                }
                //before starting the batch we need to double check some doors have been seleted incase the procedure found none
                sql = "SELECT top 1 ID FROM dbo.auto_batch_selected_door";
                int temp = 0;
                using (SqlCommand cmd = new SqlCommand(sql, conn))
                    temp = Convert.ToInt32(cmd.ExecuteScalar());
                if (temp == 0)
                {
                    Console.WriteLine("No doors able to batch... press any key to exit");
                    Console.ReadLine();
                    Environment.Exit(-1);
                }
                startBatch();
                Console.WriteLine("End of batch...");
                Console.ReadLine(); //pauses the app
                Environment.Exit(-1);
            }
        }


        private static void startBatch()
        {
            //at this point we can assume that the  auto_batch master has dumped some info into [auto_batch_selected_door]
            //and if it hasnt then end the loop entirely 
            string sql = "SELECT DISTINCT door_id FROM dbo.auto_batch_selected_door"; //fill a datatable     needs to be unique 
            int count = 0;
            DataTable distinctDoorIDDT = new DataTable();
            using (SqlConnection connCount = new SqlConnection(CONNECT.ConnectionString))
            {
                connCount.Open();
                using (SqlCommand cmdCheckData = new SqlCommand(sql, connCount))
                {
                    //NEED TO COUNT THE ROWS OF THE DATATABLE BECAUSE COUNTING DOORS IS SCUFFED
                    SqlDataAdapter da = new SqlDataAdapter(cmdCheckData);
                    da.Fill(distinctDoorIDDT);
                    count = distinctDoorIDDT.Rows.Count;
                    if (count < 1)
                        return;
                }
                connCount.Close();
            }

            //this is where things start getting giga confusing....
            //ok so i think the first step is inserting into dbo.batch


            DataTable dt = new DataTable();
            using (SqlConnection conn = new SqlConnection(CONNECT.ConnectionString))
            {
                int temp = 0;
                while (count > 0)
                {
                    conn.Open();

                    sql = "SELECT DISTINCT * FROM dbo.auto_batch_selected_door WHERE door_id = " + distinctDoorIDDT.Rows[temp][0].ToString();
                    using (SqlCommand cmdDT = new SqlCommand(sql, conn))
                    {
                        SqlDataAdapter da = new SqlDataAdapter(cmdDT);
                        da.Fill(dt);
                    }
                    //////////////////////////////////////
                    ///
                    Console.WriteLine("------------------------------------------------------------------------------------------");
                    string columns = "";
                    foreach (DataColumn col in dt.Columns)
                        columns = columns + col.ColumnName.ToString() + " -- ";
                    Console.WriteLine(columns);
                    string rowData = "";
                    foreach (DataRow row in dt.Rows)
                    {
                        foreach (DataColumn col in dt.Columns)
                        {
                            rowData = rowData + row[col].ToString() + " -- ";
                        }
                        Console.WriteLine(rowData);
                        rowData = "";
                    }
                    ///
                    /////////////////////////////////////

                    //couple of variables we need for the insert etc
                    int door_id = 0, batch_id = 0, finn = 0, rainer = 0, quantity = 0;
                    string program_no = "";  //looks like program no is only used when part batching? same with qty
                    door_id = Convert.ToInt32(distinctDoorIDDT.Rows[temp][0].ToString());
                    //program_no = dt.Rows[temp][2].ToString();
                    //quantity = Convert.ToInt32(dt.Rows[temp][7].ToString()); //work this out a lil later xoxo


                    //get max batch id rq
                    using (SqlCommand cmdBatchID = new SqlCommand("Select MAX(batch_id) + 1 FROM dbo.batch", conn))
                        batch_id = Convert.ToInt32(cmdBatchID.ExecuteScalar());
                    //set batch id to something static because we want to replicate a finished search
                    //batch_id = 9692;

                    sql = "INSERT INTO dbo.batch (door_id,batch_date,batch_id,part_batched,offcut) VALUES" +
                        "(" + door_id.ToString() + ",GETDATE()," + batch_id.ToString() + ",0,0)";
                    using (SqlCommand cmd = new SqlCommand(sql, conn))
                        cmd.ExecuteNonQuery();
                    Console.WriteLine(sql);

                    //next step in the batch is to set finn OR rainer in dbo.door
                    //Case "FINN-POWER"


                    for (int counter = 0; counter < dt.Rows.Count; counter++)
                    {
                        if (dt.Rows[counter][9].ToString() == "RAINER")
                        {
                            rainer = -1;
                            finn = 0;
                        }
                        else if (dt.Rows[counter][9].ToString() == "FINN-POWER")
                        {
                            rainer = 0;
                            finn = -1;
                        }

                        if (finn == -1)
                            sql = "UPDATE dbo.door set finn = finn + " + dt.Rows[counter][7].ToString() + " where id=" + +door_id;
                        else if (rainer == -1)
                            sql = "UPDATE dbo.door set rainer = rainer + " + dt.Rows[counter][7].ToString() + " where id= " + door_id;
                        else
                            sql = "error no machine -1";

                        using (SqlCommand cmd = new SqlCommand(sql, conn))
                            cmd.ExecuteNonQuery();
                        Console.WriteLine(sql);
                    }


                    //set as batched
                    sql = "update dbo.door set batch_live =0, batched=-1 where id =" + door_id;
                    using (SqlCommand cmd = new SqlCommand(sql, conn))
                        cmd.ExecuteNonQuery();


                    Console.WriteLine(sql);
                    count = count - 1;
                    temp = temp + 1;
                    conn.Close();
                    dt.Clear();
                }
                WriteToPunchProgQ(); //this works (breaks at the point where it wants to read from batch)header
                writeToPunchProgQRainer(); //this works  (breaks at the point where it needs to read from batch header 
                autoWriteToFinn();
            }
        }

        private static void WriteToPunchProgQ() //this is double counting
        {
            //check if there is anything for the finn(?)
            List<string> groupingList = new List<string>();   //the idea of this is that it will skip double entering when there are multiple of the same grouping :}
            Console.WriteLine("----------------------------------------------------------------------------------------------");
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
                    if (dt.Rows.Count < 1)
                        temp = 0;
                    else
                        temp = Convert.ToInt32(dt.Rows[0][0].ToString());
                }
                if (temp == 0) //need to test this 
                    return; //there is nothing for the finn

                //insert into bath header  (grab the max batch_id + 1 )

                //set batch id to something static because we want to replicate a finished search
                // batch_id = 9692;
                // batch_id = 9692;

                //loop here for EACH program number of EACH door that needs to be batched
                int doorCount = 0;
                int counter = 0;
                sql = "select count(id) from dbo.auto_batch_selected_door WHERE machine = 'FINN-POWER' ";  //this needs to check for MACHINE id only
                using (SqlCommand cmd = new SqlCommand(sql, conn))
                    doorCount = Convert.ToInt32(cmd.ExecuteScalar());
                while (counter < doorCount)
                {
                    sql = "SELECT MAX(batch_id) + 1 FROM dbo.batch";
                    using (SqlCommand cmd = new SqlCommand(sql, conn))
                        batch_id = Convert.ToInt32(cmd.ExecuteScalar());

                    //grouping
                    sql = "SELECT [group] FROM dbo.auto_batch_finn_batch where program_id = '" + dt.Rows[counter][2] + "'";
                    using (SqlCommand cmd = new SqlCommand(sql, conn))
                        grouping = Convert.ToString(cmd.ExecuteScalar());

                    sql = "SELECT [batch_id] FROM dbo.auto_batch_finn_batch where program_id = '" + dt.Rows[counter][2] + "'";
                    using (SqlCommand cmd = new SqlCommand(sql, conn))
                        batch_id = Convert.ToInt32(cmd.ExecuteScalar());  //not sure about this one  -- should work as intended now tho....



                    //////////////////////////////////////////////////////////////////

  
                   

                    if (groupingList.Contains(grouping)) //we've already run this grouping through (batch_id should keep this unique so we dont miss anything 
                    {
                        counter = counter + 1;
                        finnBatchDT.Clear();
                        continue;
                    }
                    else
                    {
                        groupingList.Add(grouping); //dont enter this path again


                        //batchheader
                        sql = "insert into dbo.batch_header (qid, qname, datecreated, machine) values ('" + batch_id + "','" + grouping + "','" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "','Finn-Power');";
                        using (SqlCommand cmd = new SqlCommand(sql, conn))
                            cmd.ExecuteNonQuery();
                        Console.WriteLine(sql);
                        Console.WriteLine("--"); ;

                        //this part does not work... pretty sure everything matches fine tho
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

                        //////////////////////////////////////
                        ///
                        Console.WriteLine("------------------------------------------------------------------------------------------");
                        string columns = "";
                        foreach (DataColumn col in finnBatchDT.Columns)
                            columns = columns + col.ColumnName.ToString() + " -- ";
                        Console.WriteLine(columns);
                        string rowData = "";
                        foreach (DataRow row in finnBatchDT.Rows)
                        {
                            foreach (DataColumn col in finnBatchDT.Columns)
                            {
                                rowData = rowData + row[col].ToString() + " -- ";
                            }
                            Console.WriteLine(rowData);
                            rowData = "";
                        }
                        ///
                        /////////////////////////////////////
                        //Console.WriteLine(sql);  //printing this sql is pretty pointless
                        Console.WriteLine("--");
                        int programCounter = 0;
                        foreach (DataRow row in finnBatchDT.Rows)
                        {
                            sql = "insert into dbo.batch_programs (door_id, program_no, sheet_quantity, header_id) values ('" + finnBatchDT.Rows[programCounter][5].ToString() + "','" + finnBatchDT.Rows[programCounter][3].ToString() + "','" + finnBatchDT.Rows[programCounter][4].ToString() + "'," + finnBatchDT.Rows[programCounter][0].ToString() + ");";
                            using (SqlCommand cmd = new SqlCommand(sql, conn))
                                cmd.ExecuteNonQuery();
                            programCounter++;
                        }

                        //Console.WriteLine(sql);
                        //Console.WriteLine("--"); ;

                    }
                    counter = counter + 1;
                    finnBatchDT.Clear();
                    // Console.WriteLine("End of WriteToPunchProgQ loop - press any key to continue");
                    //Console.ReadLine(); //pause
                }
                conn.Close();
            }
        }

        private static void writeToPunchProgQRainer()
        {
            ////////////////////////////////////////////////
            //xml variables
            List<string> groupingList = new List<string>(); //this is used for 
            string PathName = "";
            string progName = "";
            int FinishedNumber = 0;
            int TargetNumber = 0;
            string StartMode = "";
            string Sheet = "";
            int SheetMNUM = 0;
            string SheetM = "";
            string SheetX = "";
            int SheetY = 0;
            double SheetT = 0;
            int From = 0;
            int index = 0;
            string sql = "";
            int temp = 0;
            ///////////////////////////////////////////////

            using (SqlConnection conn = new SqlConnection(CONNECT.ConnectionString)) //so it doesnt create a xml file before checking if it
            {
                sql = "Select * FROM dbo.auto_batch_selected_door WHERE machine = 'RAINER'";
                DataTable dt = new DataTable();
                using (SqlCommand cmd = new SqlCommand(sql, conn))
                {
                    SqlDataAdapter da = new SqlDataAdapter(cmd);
                    da.Fill(dt);
                    temp = Convert.ToInt32(dt.Rows[0][0].ToString());
                }
                if (temp == 0)
                    return; //there is nothing for the finn

            }

            //i dont know if it needs to look at whats there but for now im just going to make a new textfile and add to that each time
            temp = 0;

            //check if there is anything for the finn(?)
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
                    batch_id = Convert.ToInt32(cmd.ExecuteScalar());

                //set batch id to something static because we want to replicate a finished search
                //batch_id = 9692;

                //loop here for EACH program number of EACH door that needs to be batched
                int doorCount = 0;
                int counter = 0;
                sql = "select count(id) from dbo.auto_batch_selected_door WHERE machine = 'RAINER'";
                using (SqlCommand cmd = new SqlCommand(sql, conn))
                    doorCount = Convert.ToInt32(cmd.ExecuteScalar());
                while (counter < doorCount)
                {
                    //grouping
                    sql = "SELECT [group] FROM dbo.auto_batch_rainer_batch where program_id = '" + dt.Rows[counter][2] + "'";
                    using (SqlCommand cmd = new SqlCommand(sql, conn))
                        grouping = Convert.ToString(cmd.ExecuteScalar());
                    //batchheader
                    sql = "insert into dbo.batch_header (qid, qname, datecreated, machine) values ('" + batch_id + "','" + grouping.ToString() + "','" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "','RAINER');";
                    if (grouping.Length < 2)
                    {
                        counter = counter + 1;
                        continue;
                    }

                    if (groupingList.Contains(grouping)) //we've already run this grouping through (batch_id should keep this unique so we dont miss anything 
                    {
                        counter = counter + 1;
                        finnBatchDT.Clear();
                        continue;
                    }
                    else
                    {
                        groupingList.Add(grouping); //dont enter this path again



                        string path = @"\\DESIGNSVR1\dropbox\xml\" + grouping + @".xml";//@"\\YWSKPC\JobList\temp_xml.txt"; //change the name to the grouping
                    if (!File.Exists(path))
                    {
                        using (StreamWriter sw = File.CreateText(path))
                        {
                            sw.WriteLine("<?xml version= \"1.0\" encoding=\"utf-8\" standalone=\"yes\"?>");
                            sw.WriteLine("<JobList>");
                            sw.Close();
                        }
                    }
                        using (var writer = new StreamWriter(path, true))
                        {

                            Console.WriteLine(sql);
                            Console.WriteLine("--");
                            using (SqlCommand cmd = new SqlCommand(sql, conn))
                                cmd.ExecuteNonQuery();

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
                            //fullSheetName = rs_rainer!group
                            string LString = "";
                            string[] LArray;

                            LString = grouping;
                            LArray = LString.Split(' ');    //Split(LString);
                            Console.WriteLine(LArray[1].ToString());
                            Console.WriteLine(LArray[2].ToString());
                            Console.WriteLine(LArray[3].ToString());
                            Console.WriteLine(LArray[4].ToString()); //test the array and see where the problem is

                            Sheet = "(0) " + LArray[2].ToString() + " " + LArray[3] + " " + LArray[4] + " " + LArray[1]; //WHAT IS THE 1 AT THE END???
                            SheetMNUM = 0;
                            SheetM = LArray[2] + " " + LArray[1];
                            SheetX = LArray[3];
                            SheetY = Convert.ToInt32(LArray[4]);
                            SheetT = Convert.ToDouble(LArray[1]);
                            From = 0;

                            //Console.WriteLine(sql); -
                            Console.WriteLine("--");


                            // 12321  -counter
                            // Console.WriteLine(finnBatchDT.Rows[counter][0].ToString());
                            //Console.WriteLine(finnBatchDT.Rows[counter][1].ToString());
                            //Console.WriteLine(finnBatchDT.Rows[counter][2].ToString());
                            //Console.WriteLine(finnBatchDT.Rows[counter][3].ToString());
                            //Console.WriteLine(finnBatchDT.Rows[counter][4].ToString());
                            //Console.WriteLine(finnBatchDT.Rows[counter][5].ToString());
                            index = 0;
                            for (int row = 0; row < finnBatchDT.Rows.Count; row++)
                            {
                                sql = "insert into dbo.batch_programs (door_id, program_no, sheet_quantity, header_id) values ('" + finnBatchDT.Rows[row][5].ToString() + "','" + finnBatchDT.Rows[row][3].ToString() + "','" + finnBatchDT.Rows[row][4].ToString() + "'," + finnBatchDT.Rows[row][0].ToString() + ");";
                                using (SqlCommand cmd = new SqlCommand(sql, conn))
                                {
                                    cmd.ExecuteNonQuery();
                                }
                                writer.Write("<Job>" + Environment.NewLine);
                                writer.Write("<Index>" + index.ToString() + "</Index>" + Environment.NewLine); ; //+ 1 index each time :}
                                writer.Write(@"<PathName>Z:\" + finnBatchDT.Rows[row][3].ToString() + ".MPF</PathName>" + Environment.NewLine);  //need to check the file type here 
                                writer.Write("<Name>" + finnBatchDT.Rows[row][3].ToString() + ".MPF</Name>" + Environment.NewLine);
                                writer.Write("<FinishedNumber>0</FinishedNumber>" + Environment.NewLine);  //always 0 ??
                                writer.Write("<TargetNumber>" + TargetNumber + "</TargetNumber>" + Environment.NewLine);
                                writer.Write("<StartMode>A</StartMode>" + Environment.NewLine); //always A?
                                writer.Write("<Sheet>" + Sheet + "</Sheet>" + Environment.NewLine);
                                writer.Write("<SheetMNUM>" + SheetMNUM + "</SheetMNUM>" + Environment.NewLine);
                                writer.Write("<SheetM>" + SheetM + "</SheetM>" + Environment.NewLine);
                                writer.Write("<SheetX>" + SheetX + "</SheetX>" + Environment.NewLine);
                                writer.Write("<SheetY>" + SheetY + "</SheetY>" + Environment.NewLine);
                                writer.Write("<SheetT>" + SheetT + "</SheetT>" + Environment.NewLine);  //these need some adjusting like the grouping will get scuffed if it has an extra space (think this has been fixed automatically by having the material set to zintex/gav statically)
                                writer.Write("<From>" + From + "</From>" + Environment.NewLine);          //needs to be tested tho
                                writer.Write("</Job>" + Environment.NewLine);
                                index = index + 1;
                                row = row++;
                            }
                            //Console.WriteLine(sql);
                            //Console.WriteLine("--");
                            counter = counter + 1;
                            Console.WriteLine("End of WriteToPunchProgQRAINER loop - press any key to continue");
                            Console.ReadLine(); //pause
                                                //write to the xml file here 



                            writer.Write("</JobList>");
                        }
                    }
                   
                }
                conn.Close();
            }
        }
        

        private static void autoWriteToFinn() //this wrtes to the table thats on the finn - just need to select whats in selected door
        {
            string sql = "";
            DataTable dtBatchIdList = new DataTable();
            DataTable dtDoorList = new DataTable();
            using (SqlConnection conn = new SqlConnection(CONNECT.ConnectionString))
            {
                conn.Open();
                int temp = 0;
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

                sql = "select distinct max(batch_id) as batch_id,max(door_id) as door_id,max(program_id) as program_id,max(FirstOfMaterial_type) as FirstOfMaterial_type," +
                    "max(FirstOfthickness) as FirstOfthickness,max(FirstOflength) as FirstOflength,max(FirstOfwidth) as FirstOfwidth,max(FirstOfquantity) as FirstOfquantity,max(FirstOfmachine) as FirstOfmachine," +
                    "max([group]) as [group],max(batch_date) as batch_date,max(part_batched) as part_batched,max(offcut) as offcut " +
                    "from dbo.auto_batch_finn_batch where FirstOfmachine = 'FINN-POWER' group by batch_id "; //loop through each of the doors (then loop through each of the program_ids on that foor
                using (SqlCommand cmd = new SqlCommand(sql, conn))
                {
                    SqlDataAdapter da = new SqlDataAdapter(cmd);
                    da.Fill(dtDoorList);
                }
                for (int rowCount = 0; rowCount < dtDoorList.Rows.Count; rowCount++)
                {
                    Console.WriteLine("-------------------------------------------------");
                    int queue = 0;
                    OleDbConnection connection = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + @"\\designsvr1\DropBox\production.mdb");//@"\\fujitsu\fp\production.mdb");
                    connection.Open();
                    OleDbDataReader reader = null;
                    OleDbCommand command = new OleDbCommand("SELECT MAX (QUEUEID) FROM QUEUE ", connection);
                    reader = command.ExecuteReader();
                    while (reader.Read())
                    {
                        queue = Convert.ToInt32(reader[0].ToString());
                        queue = queue + 1;
                    }

                    //insert into queue
                    sql = "insert into QUEUE (QUEUEID, QNAME) values ('" + queue + "','" + dtDoorList.Rows[rowCount][9] + "')";
                    OleDbCommand cmd2 = new OleDbCommand(sql, connection);
                    cmd2.ExecuteNonQuery();


                    sql = "select distinct * from dbo.auto_batch_finn_batch where FirstOfmachine = 'FINN-POWER' AND door_id = " + dtDoorList.Rows[rowCount][1].ToString() + " order by program_id ";
                    using (SqlCommand cmd = new SqlCommand(sql, conn))
                    {
                        SqlDataAdapter da = new SqlDataAdapter(cmd);
                        da.Fill(dtBatchIdList);
                    }

                    //now we loop through the DT
                    //needs to have access to the finn

                    //
                    int progressQNumber = 0;
                    for (int row = 0; row < dtBatchIdList.Rows.Count; row++)
                    {
                        sql = "INSERT into QUEUEPROG (QUEUEID, PROGSEQ, NCFILE, TOTALAMOUNT, THICKNESS) VALUES ('" + queue + "','" + progressQNumber.ToString() + "','" + @"Z:\" + dtBatchIdList.Rows[row][2].ToString() + ".NC" + "','" + dtBatchIdList.Rows[row][7].ToString() + "','" + dtBatchIdList.Rows[row][4].ToString() + "');";
                        OleDbCommand cmd = new OleDbCommand(sql, connection);
                        cmd.ExecuteNonQuery();
                        progressQNumber = progressQNumber + 1;
                        Console.WriteLine(sql);

                        //queue = queue + 1;
                    }
                    dtBatchIdList.Clear();
                }

                    /*//////////////////////////////////////////////// 

                    DoCmd.SetWarnings False
                    Do While rs.EOF = False

                        QUEUEID = DMax("QUEUEID", "QUEUE") 
                        QUEUEID = QUEUEID + 1


                        'insert group into table
                        sql = "insert into QUEUE (QUEUEID, QNAME) values ('" & QUEUEID & "','" & rs!batch_id & RemoveFullStop(rs!FirstOfFirstOfthickness) & rs!FirstOfFirstOfmaterial_type & rs!FirstOfFirstOflength & rs!FirstOfFirstOfWidth & "');"
                        DoCmd.RunSQL sql


                            'select for programs based on the group
                            sql_select2 = "SELECT QUEUE.QUEUEID, QUEUE.QNAME, RemoveFullStop2(CStr([FirstOfthickness])) AS format_thickness, [batch_id] & [format_thickness] & [FirstOfmaterial_type] & [FirstOflength] & [FirstOfwidth] AS group2, qryFinnQNameFormat.FirstOfmaterial_type, qryFinnQNameFormat.FirstOfthickness, qryFinnQNameFormat.FirstOflength, qryFinnQNameFormat.FirstOfwidth, qryFinnQNameFormat.batch_id, qryFinnQNameFormat.program_no, qryFinnQNameFormat.FirstOfquantity " & _
                                "FROM qryFinnQNameFormat INNER JOIN QUEUE ON qryFinnQNameFormat.group2 = QUEUE.QNAME GROUP BY QUEUE.QUEUEID, QUEUE.QNAME, RemoveFullStop2(CStr([FirstOfthickness])), qryFinnQNameFormat.FirstOfmaterial_type, qryFinnQNameFormat.FirstOfthickness, qryFinnQNameFormat.FirstOflength, qryFinnQNameFormat.FirstOfwidth, qryFinnQNameFormat.batch_id, qryFinnQNameFormat.program_no, qryFinnQNameFormat.FirstOfquantity HAVING (((QUEUE.QNAME)='" & rs!batch_id & RemoveFullStop(rs!FirstOfFirstOfthickness) & rs!FirstOfFirstOfmaterial_type & rs!FirstOfFirstOflength & rs!FirstOfFirstOfWidth & "'));"


                            Set db = CurrentDb
                            Set rs2 = db.OpenRecordset(sql_select2, dbOpenDynaset)


                            PROGSEQNO = 0
                            rs2.MoveFirst

                            Do While rs2.EOF = False
                                PROGSEQNO = PROGSEQNO + 1
                               'inserts program_no into QUEUEPROG
                               sql_insert = "INSERT into QUEUEPROG (QUEUEID, PROGSEQ, NCFILE, TOTALAMOUNT, THICKNESS) VALUES ('" & QUEUEID & "','" & PROGSEQNO & "','" & "Z:\" & rs2!program_no & ".NC" & "','" & rs2!FirstOfquantity & "','" & rs2!FirstOfthickness & "');"
                               DoCmd.RunSQL sql_insert
                               rs2.MoveNext
                           Loop


                        rs.MoveNext

                    Loop

                    using (sqlcommand cmd cmd -new now fdshdsk)
                    DoCmd.SetWarnings True


                    
                    *////////////////////////////////////////////////
                    conn.Close();
            }
        }

    }
}

