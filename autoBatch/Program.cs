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
        public static int rotecSGBatchID { get; set; }
        public static int rotecNonSGBatchID { get; set; }
        public static int SGBatchID { get; set; }
        public static int NonSGBatchID { get; set; }
        static void Main(string[] args)
        {
            int bufferHeight = Console.BufferHeight;
            int bufferWidth = Console.BufferWidth;
            string sql = "";


            //bufferHeight += 240;
            //Console.BufferHeight = bufferHeight;

            bufferWidth += 85;
            Console.BufferWidth = bufferWidth; //make the console app resizable  (bigger)

            //first thing to handle is to check if there is enough doors to batch at all!
            using (SqlConnection conn = new SqlConnection(CONNECT.ConnectionString))
            {
                conn.Open();
                int minDoors = 0;
                int currentDoors = 0;
                sql = "SELECT minimum_doors from dbo.auto_batch_limit";
                using (SqlCommand cmd = new SqlCommand(sql, conn))
                    minDoors = Convert.ToInt32(cmd.ExecuteScalar());
                sql = "select count(a.id) as [count] from dbo.door a LEFT JOIN dbo.door_program b ON a.id = b.door_id LEFT JOIN dbo.door_type c ON a.door_type_id = c.id " +
                    "WHERE b.checked_by_id is not null AND batched<> -1 AND date_punch > '2020-06-01' AND(status_id = 1 or status_id = 2) AND(c.id <> 43 OR c.id <> 11  OR c.id <> 123  OR c.id <> 124  OR c.id <> 125  OR c.id <> 140  OR c.id <> 150  OR c.id <> 151)";
                using (SqlCommand cmd = new SqlCommand(sql, conn))
                    currentDoors = Convert.ToInt32(cmd.ExecuteScalar());

                conn.Close();
                if (currentDoors < minDoors) //theres not enough doors to batch  so exit
                    return;
            }

            //autoWriteToFinn();
            //return;
            string path = @"\\DESIGNSVR1\subcontracts\Express5.MAC\FINNPRODUCTION.csv";  //@"\\DESIGNSVR1\dropbox\FINNPRODUCTION_temp.csv";  //
            Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(path);
            Microsoft.Office.Interop.Excel.Worksheet xlWorksheet = xlWorkbook.Sheets[1]; // assume it is the first sheet
            Microsoft.Office.Interop.Excel.Range xlRange = xlWorksheet.UsedRange; // get the entire used range
            int value = 10;
            int numberOfColumnsToRead = value;
            int last = xlWorksheet.Cells.SpecialCells(Microsoft.Office.Interop.Excel.XlCellType.xlCellTypeLastCell, Type.Missing).Row;
            Microsoft.Office.Interop.Excel.Range range = xlWorksheet.get_Range("A1:A" + last);
            
            using (SqlConnection conn = new SqlConnection(CONNECT.ConnectionString))
            {
                conn.Open();
                //delete entried in that table here?
                sql = "DELETE FROM dbo.auto_batch_finn_csv_import ";
                using (SqlCommand cmd = new SqlCommand(sql, conn))
                    cmd.ExecuteNonQuery();
                sql = "DELETE FROM dbo.auto_batch_selected_door"; //wipe the old selected list here too! otherwise we won't get anywhere~
                using (SqlCommand cmd = new SqlCommand(sql, conn))
                    cmd.ExecuteNonQuery();
                    for (int i = 1; i < 500; i++)  //for (int i = 1; i < last; i++)      //for X amount of rows in the excel sheet
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
                        tempString = tempString.Trim();  //remove the leading spaces in material (for some reason there is white space at start of it in the csv by default

                        sql = "INSERT INTO dbo.auto_batch_finn_csv_import (door_id,program_id,material,thickness,length,width,quantity,date,machine) VALUES (" +
                       temp + ",'" + //door id
                       xlRange.Cells[i, 2].Value2.ToString() + "','" +//program id
                       tempString + "'," +//material (with absouletly no leading or trailing white spaces!!! :DDD 
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
            xlApp.Quit(); //this isnt enough to properly close the app... 

            //this will loop through every process and kill anything that is related to excel - this is probably fine as it'll be run somewhere where there is no user opening excel files
            System.Diagnostics.Process[] process = System.Diagnostics.Process.GetProcessesByName("Excel");
            foreach (System.Diagnostics.Process p in process)
            {
                if (!string.IsNullOrEmpty(p.ProcessName))
                {
                    try
                    {
                        p.Kill(); //kills each process :}
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
                if (current_number + 1 < limit) //i think this section needs to add one onto the current number
                {
                    while (current_number + 1 < limit)
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
                    Environment.Exit(-1); // exit out of the app
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
                //moveXML();
                //printSheets();
                Console.WriteLine("End of batch...");
                Console.ReadLine(); //pauses the app
                Environment.Exit(-1);
            }
        }


        private static void startBatch()
        {
            //at this point we can assume that the  auto_batch master has dumped some info into [auto_batch_selected_door]
            //and if it hasnt then exit
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

            //insert 


            DataTable dt = new DataTable();
            using (SqlConnection conn = new SqlConnection(CONNECT.ConnectionString))
            {
                conn.Open();
                //here we set the batchIDs of the  each TYPE of  batch
                int batchTemp = 0;
                sql = "SELECT MAX(batch_id) + 1 FROM dbo.batch";
                using (SqlCommand cmd = new SqlCommand(sql, conn))
                    batchTemp = Convert.ToInt32(cmd.ExecuteScalar());
                rotecSGBatchID = batchTemp;
                rotecNonSGBatchID = batchTemp + 1;  //add one per type so we get a unique id for all of these, although this will prob bug out if another person batches at the exact same time -- should probably adjust for this ~
                SGBatchID = batchTemp + 2;
                NonSGBatchID = batchTemp + 3;

                int temp = 0;
                while (count > 0)
                {
                    sql = "SELECT DISTINCT * FROM dbo.auto_batch_selected_door WHERE door_id = " + distinctDoorIDDT.Rows[temp][0].ToString() + " Order by program_id asc";
                    using (SqlCommand cmdDT = new SqlCommand(sql, conn))
                    {
                        SqlDataAdapter da = new SqlDataAdapter(cmdDT);
                        da.Fill(dt);
                    }
                    //////////////////////////////////////   //this is purely to read the datatable
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
                    /////////////////////////////////////

                    //couple of variables we need for the insert etc
                    int door_id = 0, batch_id = 0, finn = 0, rainer = 0, quantity = 0;
                    string program_no = "";
                    door_id = Convert.ToInt32(distinctDoorIDDT.Rows[temp][0].ToString());
                    //program_no = dt.Rows[temp][2].ToString();


                    //get max batch id rq
                    using (SqlCommand cmdBatchID = new SqlCommand("Select MAX(batch_id) + 1 FROM dbo.batch", conn))
                        batch_id = Convert.ToInt32(cmdBatchID.ExecuteScalar());

                    if (dt.Rows[0][10].ToString() == "SG")
                        batch_id = SGBatchID;
                    else if (dt.Rows[0][10].ToString() == "NonSG")
                        batch_id = NonSGBatchID;
                    else if (dt.Rows[0][10].ToString() == "RotecNonSG")
                        batch_id = rotecNonSGBatchID;
                    else if (dt.Rows[0][10].ToString() == "RotecSG")
                        batch_id = rotecSGBatchID;


                    sql = "INSERT INTO dbo.batch (door_id,batch_date,batch_id,part_batched,offcut) VALUES" +
                    "(" + door_id.ToString() + ",GETDATE()," + batch_id.ToString() + ",0,0)";
                    using (SqlCommand cmd = new SqlCommand(sql, conn))
                        cmd.ExecuteNonQuery();
                    Console.WriteLine(sql);

                    //next step in the batch is to set finn OR rainer in dbo.door



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
                    dt.Clear();
                }
                conn.Close();
                WriteToPunchProgQ();
                writeToPunchProgQRainer();
                autoWriteToFinn();
                //printSheets();  12321
            }
        }

        private static void WriteToPunchProgQ()
        {
            //check if there is anything for the finn
            List<string> groupingList = new List<string>();   //add each grouping to the list so we can check for doubles (each loop adds for every line item so we only need to run once per grouping)
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
                if (temp == 0)
                    return; //there is nothing for the finn

                //insert into bath header  (grab the max batch_id + 1 )

                //loop here for EACH program number of EACH door that needs to be batched
                int doorCount = 0;
                int counter = 0;
                sql = "select count(id) from dbo.auto_batch_selected_door WHERE machine = 'FINN-POWER' ";
                using (SqlCommand cmd = new SqlCommand(sql, conn))
                    doorCount = Convert.ToInt32(cmd.ExecuteScalar());
                while (counter < doorCount)
                {
                    sql = "SELECT MAX(batch_id) + 1 FROM dbo.batch";
                    using (SqlCommand cmd = new SqlCommand(sql, conn))
                        batch_id = Convert.ToInt32(cmd.ExecuteScalar());
                    //^^ this needs to be removed  but for now i just overwrite it below~


                    //grouping
                    sql = "SELECT [group] FROM dbo.auto_batch_finn_batch where program_id = '" + dt.Rows[counter][2] + "'";
                    using (SqlCommand cmd = new SqlCommand(sql, conn))
                        grouping = Convert.ToString(cmd.ExecuteScalar());

                    //batch_id
                    sql = "SELECT [batch_id] FROM dbo.auto_batch_finn_batch where program_id = '" + dt.Rows[counter][2] + "'";
                    using (SqlCommand cmd = new SqlCommand(sql, conn))
                        batch_id = Convert.ToInt32(cmd.ExecuteScalar());



                    //////////////////////////////////////////////////////////////////




                    if (groupingList.Contains(grouping)) //we've already run this grouping through (batch_id should keep this unique so we dont miss anything 
                    {
                        counter = counter + 1;
                        finnBatchDT.Clear();
                        continue;
                    }
                    else
                    {
                        groupingList.Add(grouping); //dont enter this path again for this grouping


                        //batchheader
                        sql = "insert into dbo.batch_header (qid, qname, datecreated, machine) values ('" + batch_id + "','" + grouping + "','" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "','Finn-Power');";
                        using (SqlCommand cmd = new SqlCommand(sql, conn))
                            cmd.ExecuteNonQuery();
                        Console.WriteLine(sql);
                        Console.WriteLine("--"); ;



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

                        //////////////////////////////////////               //this is only to output what is in the datatable right now
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

        private static void writeToPunchProgQRainer() //is actuall yawei
        {
            ////////////////////////////////////////////////
            //xml variables
            List<string> groupingList = new List<string>(); //this is used for stopping double entries, same as above
            string PathName = ""; //most of these are not used
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

            using (SqlConnection conn = new SqlConnection(CONNECT.ConnectionString)) //check if there are any doors here  and if not exit out (to stop empty .xml files being made
            {
                sql = "Select * FROM dbo.auto_batch_selected_door WHERE machine = 'RAINER'";
                DataTable dt = new DataTable();
                using (SqlCommand cmd = new SqlCommand(sql, conn))
                {
                    SqlDataAdapter da = new SqlDataAdapter(cmd);
                    da.Fill(dt);
                    if (dt.Rows.Count > 1)
                        temp = Convert.ToInt32(dt.Rows[0][0].ToString());
                    else temp = 0;
                }
                if (temp == 0)
                    return; //there is nothing for the finn
            }


            temp = 0;

            //check if there is anything for the rainer
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
                    if (dt.Rows.Count > 1)
                        temp = Convert.ToInt32(dt.Rows[0][0].ToString());
                    else
                        temp = 0;
                }
                if (temp == 0)
                    return; //there is nothing for the finn

                //insert into bath header  (grab the max batch_id + 1 )

                sql = "SELECT MAX(batch_id) + 1 FROM dbo.batch";
                using (SqlCommand cmd = new SqlCommand(sql, conn))
                    batch_id = Convert.ToInt32(cmd.ExecuteScalar());


                //12321

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
                    if (dt.Rows[0][10].ToString() == "SG")
                        batch_id = SGBatchID;
                    else if (dt.Rows[0][10].ToString() == "NonSG")
                        batch_id = NonSGBatchID;
                    else if (dt.Rows[0][10].ToString() == "RotecNonSG")
                        batch_id = rotecNonSGBatchID;
                    else if (dt.Rows[0][10].ToString() == "RotecSG")
                        batch_id = rotecSGBatchID;
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
                        groupingList.Add(grouping); //dont enter this path again for this grouping 

                        string path =  @"\\YWSKPC\JobList\" + grouping + @".xml"; //@"//DESIGNSVR1\dropbox\xml\" + grouping + @".xml";  //
                        try
                        {
                            if (!File.Exists(path))
                            {
                                using (StreamWriter sw = File.CreateText(path))
                                {
                                    sw.WriteLine("<?xml version= \"1.0\" encoding=\"utf-8\" standalone=\"yes\"?>");
                                    sw.WriteLine("<JobList>");
                                    sw.Close();
                                }
                            }
                        }
                        catch
                        {
                            path = @"//DESIGNSVR1\dropbox\IT_FAILED\XML\" + grouping + @".xml";
                            if (!File.Exists(path))
                            {
                                using (StreamWriter sw = File.CreateText(path))
                                {
                                    sw.WriteLine("<?xml version= \"1.0\" encoding=\"utf-8\" standalone=\"yes\"?>");
                                    sw.WriteLine("<JobList>");
                                    sw.Close();
                                }
                            }
                        }

                        using (var writer = new StreamWriter(path, true))
                        {

                            Console.WriteLine(sql);
                            Console.WriteLine("--");
                            using (SqlCommand cmd = new SqlCommand(sql, conn))
                                cmd.ExecuteNonQuery();


                            sql = "SELECT dbo.batch_header.id, dbo.batch_header.qid, dbo.batch_header.qname, auto_batch_rainer_batch.program_id, auto_batch_rainer_batch.FirstOfquantity, auto_batch_rainer_batch.door_id " +
                                "FROM auto_batch_rainer_batch " +
                                "INNER JOIN dbo.batch_header ON auto_batch_rainer_batch.[group] = dbo.batch_header.qname " +
                                "WHERE [group] = '" + grouping + "' AND machine = 'rainer' " +
                                "GROUP BY dbo.batch_header.id, dbo.batch_header.qid, dbo.batch_header.qname, auto_batch_rainer_batch.program_id, auto_batch_rainer_batch.FirstOfquantity, auto_batch_rainer_batch.door_id";
                            using (SqlCommand cmd = new SqlCommand(sql, conn))
                            {
                                SqlDataAdapter da = new SqlDataAdapter(cmd);
                                da.Fill(finnBatchDT);
                            }
                            //fullSheetName = rs_rainer!group

                            /////////////////////////////////////////////////////////////////
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
                            /////////////////////////////////////////////////////////////////
                            string LString = "";
                            string[] LArray;

                            LString = grouping;
                            LArray = LString.Split(' ');    //Split(LString);


                            //fill variables from the grouping
                            Sheet = "(0) " + LArray[2].ToString() + " " + LArray[3] + " " + LArray[4] + " " + LArray[1]; //WHAT IS THE 1 AT THE END???
                            SheetMNUM = 0;
                            SheetM = LArray[2] + " " + LArray[1];
                            SheetX = LArray[3];
                            SheetY = Convert.ToInt32(LArray[4]);
                            SheetT = Convert.ToDouble(LArray[1]);
                            From = 0;

                            //Console.WriteLine(sql); -
                            Console.WriteLine("--");



                            // Console.WriteLine(finnBatchDT.Rows[counter][0].ToString());
                            //Console.WriteLine(finnBatchDT.Rows[counter][1].ToString());
                            //Console.WriteLine(finnBatchDT.Rows[counter][2].ToString());
                            //Console.WriteLine(finnBatchDT.Rows[counter][3].ToString());
                            //Console.WriteLine(finnBatchDT.Rows[counter][4].ToString());
                            //Console.WriteLine(finnBatchDT.Rows[counter][5].ToString());
                            index = 0;
                            for (int row = 0; row < finnBatchDT.Rows.Count; row++) //add each item to the xml file, loops for each item that has the same grouping
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
                                writer.Write("<TargetNumber>" + finnBatchDT.Rows[row][4].ToString() + "</TargetNumber>" + Environment.NewLine);
                                writer.Write("<StartMode>A</StartMode>" + Environment.NewLine); //always A?
                                writer.Write("<Sheet>" + Sheet + "</Sheet>" + Environment.NewLine);
                                writer.Write("<SheetMNUM>" + SheetMNUM + "</SheetMNUM>" + Environment.NewLine);
                                writer.Write("<SheetM>" + SheetM + "</SheetM>" + Environment.NewLine);
                                writer.Write("<SheetX>" + SheetX + "</SheetX>" + Environment.NewLine);
                                writer.Write("<SheetY>" + SheetY + "</SheetY>" + Environment.NewLine);
                                writer.Write("<SheetT>" + SheetT + "</SheetT>" + Environment.NewLine);
                                writer.Write("<From>" + From + "</From>" + Environment.NewLine);
                                writer.Write("</Job>" + Environment.NewLine);
                                index = index + 1;
                                row = row++;
                            }
                            //Console.WriteLine(sql);
                            //Console.WriteLine("--");
                            counter = counter + 1;
                            Console.WriteLine("--------");
                            //Console.WriteLine("End of WriteToPunchProgQRAINER loop - press any key to continue");
                            //  Console.ReadLine(); //pause




                            writer.Write("</JobList>"); //close up this xml file
                        }
                    }

                }
                conn.Close();
            }
        }


        private static void autoWriteToFinn() //this wrtes to the table thats on the finn -
        {

            string sql = "";
            DataTable dtDoorList = new DataTable();
            DataTable dtGroupList = new DataTable();
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
                    if (dt.Rows.Count > 1)
                        temp = Convert.ToInt32(dt.Rows[0][0].ToString());
                    else temp = 0;
                }
                if (temp == 0)
                    return; //there is nothing for the finn

                sql = "select[group] from dbo.auto_batch_finn_batch where FirstOfmachine = 'FINN-POWER' group by [group]"; //loop each of the groups
                using (SqlCommand cmd = new SqlCommand(sql, conn))
                {
                    SqlDataAdapter da = new SqlDataAdapter(cmd);
                    da.Fill(dtGroupList);
                }

                /////////////////////////////////////////////////////////////////
                string rowData = "";
                foreach (DataRow row in dtGroupList.Rows)
                {
                    foreach (DataColumn col in dtGroupList.Columns)
                    {
                        rowData = rowData + row[col].ToString() + " -- ";
                    }
                    Console.WriteLine(rowData);
                    rowData = "";
                }
                /////////////////////////////////////////////////////////////////



                for (int rowCount = 0; rowCount < dtGroupList.Rows.Count; rowCount++) //for every group ----
                {
                    //we now need to gather data based on the group!
                    sql = "select distinct * from dbo.auto_batch_finn_batch where FirstOfmachine = 'FINN-POWER' AND [group] = '" + dtGroupList.Rows[rowCount][0].ToString() + "' order by program_id ";
                    using (SqlCommand cmd = new SqlCommand(sql, conn))
                    {
                        SqlDataAdapter da = new SqlDataAdapter(cmd);
                        da.Fill(dtDoorList);
                    }
                    Console.WriteLine("-------------------------------------------------");
                    /////////////////////////////////////////////////////////////////
                    rowData = "";
                    foreach (DataRow row in dtDoorList.Rows)
                    {
                        foreach (DataColumn col in dtDoorList.Columns)
                        {
                            rowData = rowData + row[col].ToString() + " -- ";
                        }
                        Console.WriteLine(rowData);
                        rowData = "";
                    }
                    /////////////////////////////////////////////////////////////////


                    Console.WriteLine("-------------------------------------------------");
                    int queue = 0;
                    OleDbConnection connection = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + @"\\fujitsu\fp\production.mdb"); //writes to the file on finn //\\designsvr1\DropBox\production.mdb   // @"\\designsvr1\DropBox\production.mdb");//  \\fujitsu\fp\production.mdb
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
                    string groupTrim = dtGroupList.Rows[rowCount][0].ToString();
                    groupTrim = groupTrim.Replace(" ", "");
                    groupTrim = groupTrim.Replace(".", "");
                    sql = "insert into QUEUE (QUEUEID, QNAME) values ('" + queue + "','" + groupTrim + "')"; //unique id per grouping
                    Console.WriteLine(sql);
                    Console.WriteLine("-------------------------------------------------");
                    OleDbCommand cmd2 = new OleDbCommand(sql, connection);
                    cmd2.ExecuteNonQuery();


                    //now we loop through the DT
                    //needs to have access to the finn
                    int progressQNumber = 1;
                    for (int row = 0; row < dtDoorList.Rows.Count; row++)
                    {
                        sql = "INSERT into QUEUEPROG (QUEUEID, PROGSEQ, NCFILE, TOTALAMOUNT, THICKNESS) VALUES ('" + queue + "','" + progressQNumber.ToString() + "','" + @"Z:\" + dtDoorList.Rows[row][2].ToString() + ".NC" + "','" + dtDoorList.Rows[row][7].ToString() + "','" + dtDoorList.Rows[row][4].ToString() + "');";
                        OleDbCommand cmd = new OleDbCommand(sql, connection);
                        cmd.ExecuteNonQuery();
                        progressQNumber = progressQNumber + 1;
                        Console.WriteLine(sql);

                        //queue = queue + 1;   
                    }
                    dtDoorList.Clear();
                    connection.Close();
                    Console.WriteLine("-------------------------------------------------");
                    Console.WriteLine("-------------------------------------------------");
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


        private static void printSheets() //here we make a .bat and that will open normal batching app and from there print out what has been inserted 
        {
            //run batch file here
            //System.Diagnostics.Process.Start(@"//DESIGNSVR1\dropbox\IT_FAILED\auto_batch_print.bat");
            int exitCode;
            System.Diagnostics.ProcessStartInfo processInfo;
            System.Diagnostics.Process process;



            processInfo = new System.Diagnostics.ProcessStartInfo(@"\\DESIGNSVR1\dropbox\IT_FAILED\auto_batch_print.bat");
            processInfo.CreateNoWindow = true;
            processInfo.UseShellExecute = false;
            process = System.Diagnostics.Process.Start(processInfo);
            process.WaitForExit();
            exitCode = process.ExitCode;
            process.Close();
        }

        private static void moveXML()
        {
            //try and move old xmls here IF they exist
            string sourceDirectory = @"//DESIGNSVR1\dropbop\xml";
            string destinationDirectory = @"//DESIGNSVR1\dropbox\IT_FAILED\moved_xml";

            try
            {
                //Directory.Move(sourceDirectory, destinationDirectory);
                Console.WriteLine("XML FILES MOVED TO MAIN DIRECTORY");
                foreach (var file in Directory.EnumerateFiles(sourceDirectory))
                {
                    string destFile = Path.Combine(destinationDirectory, Path.GetFileName(file));
                    if (!File.Exists(destFile))
                        File.Move(file, destFile);

                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
        }
    }
}

