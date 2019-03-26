using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;
using System.Timers;
using System.IO;
using System.Data.SqlClient;
using System.Data.OleDb;


namespace LogService
{
    public partial class Service1 : ServiceBase
    {
       // Service1 service = new Service1();
      

        public void OnStart()
        {
            System.Diagnostics.Debugger.Launch();
            string filePath = "D:\\PTACDealerList.xlsx";

            if (System.IO.File.Exists(filePath))
            {
                string extension = Path.GetExtension(filePath);
                string conString = string.Empty;
                switch (extension)
                {
                    case ".xls": //Excel 97-03.
                        conString = System.Configuration.ConfigurationManager.ConnectionStrings["Excel03ConString"].ConnectionString;
                        
                        break;
                    case ".xlsx": //Excel 07 and above.
                        conString = System.Configuration.ConfigurationManager.ConnectionStrings["Excel07ConString"].ConnectionString;
                        break;
                }


                DataTable dt = new DataTable();

                conString = string.Format(conString, filePath);

                using (OleDbConnection connExcel = new OleDbConnection(conString))
                {
                    using (OleDbCommand cmdExcel = new OleDbCommand())
                    {
                        using (OleDbDataAdapter odaExcel = new OleDbDataAdapter())
                        {
                            cmdExcel.Connection = connExcel;

                            //Get the name of First Sheet.
                            connExcel.Open();
                            DataTable dtExcelSchema;
                            dtExcelSchema = connExcel.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                            string sheetName = dtExcelSchema.Rows[0]["TABLE_NAME"].ToString();
                            connExcel.Close();

                            //Read Data from First Sheet.
                            connExcel.Open();
                            cmdExcel.CommandText = "SELECT * From [" + sheetName + "]";
                            odaExcel.SelectCommand = cmdExcel;
                            odaExcel.Fill(dt);
                            foreach (DataRow row in dt.Rows)
                            {
                                var Description = row["SUBFDESC"].ToString().Trim();
                                var Address = row["SUBFADR1"].ToString().Trim();
                                var City = row["SUBFCITY"].ToString().Trim();
                                var State = row["SUBFSTATE"].ToString().Trim();
                                var ZipCode = row["SUBFZIP"].ToString().Trim();
                                var PhoneNumber = row["PhoneNum"].ToString().Trim();
                                var PhoneNumber2 = row["PhoneNum"].ToString().Trim();

                                conString = System.Configuration.ConfigurationManager.ConnectionStrings["constr"].ConnectionString;
                                using (SqlConnection con = new SqlConnection(conString))
                                {
                                    String query = "INSERT INTO dbo.PtacDealer_TB (Description,Address,City,State,ZipCode,PhoneNumber,PhoneNumber2) VALUES ('" + Description.Replace("'", "''") + "','" + Address.Replace("'", "''") + "','" + City.Replace("'", "''") + "','" + State + "','" + ZipCode + "','" + PhoneNumber + "','" + PhoneNumber2 + "')";

                                    SqlCommand command = new SqlCommand(query, con);

                                    //var desc =  command.Parameters.Add("@Description", Description);
                                    //  command.Parameters.Add("@Address", Address);
                                    //  command.Parameters.Add("@City", City);
                                    //  command.Parameters.Add("@State", State);
                                    //  command.Parameters.Add("@ZipCode", ZipCode);
                                    //  command.Parameters.Add("@PhoneNumber", PhoneNumber);
                                    //  command.Parameters.Add("@PhoneNumber2", PhoneNumber2);
                                    //con.Open();
                                    //string  result = command.ExecuteNonQuery();


                                    con.Open();
                                    command.ExecuteNonQuery();
                                    con.Close();
                                    //command.ExecuteNonQuery();
                                    connExcel.Close();
                                }


                            }
                        }
                    }
                }
            } 
        }

        //Timer tmr = new Timer();
        //public Service1()
        //{
        //    InitializeComponent();
        //}

        //protected override void OnStart(string[] args)
        //{
        //    tmr.Enabled = true;
        //    tmr.Interval = 1500;
        //    tmr.Elapsed += tmr_Elapsed;
        //}
       
        //void tmr_Elapsed(Object sender, ElapsedEventArgs e)
        //{
        //    if(File.Exists(@"D:/Practice/MVC Practice Programs/LogFiles/logfile.txt"))
        //    {
        //        File.AppendAllText(@"D:/Practice/MVC Practice Programs", "Log Time:" + System.DateTime.Now.ToString() + "\n");
        //    }
        //    else{
        //        File.Create(@"D:/Practice/MVC Practice Programs/LogFiles/logfile.txt");
        //        File.AppendAllText(@"D:/Practice/MVC Practice Programs", "Log Time:"+System.DateTime.Now.ToString()+"\n");

        //    }
        //}

        //protected override void OnStop()
        //{
        //    tmr.Enabled = false;
        //}
    }
}
