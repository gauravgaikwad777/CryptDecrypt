using OfficeOpenXml;
using Oracle.ManagedDataAccess.Client;
using System;
using System.Collections.Generic;
using System.IO;
using upes_admission.Gutils;

namespace export_student_excel
{
    class Program
    {
        static void Main(string[] args)
        {
            List<Student> list = GetDataTableFromDB("Select id, migrated_password, first_name, last_name, email_address, country_id from donor where migrated_password is not null and id = 82 order by id ");
            string filename = string.Format("Password-decrypt-{0:yyyy-MM-dd-HH-mm-ss}.xlsx", DateTime.UtcNow);

            //To-Do
            string rootFolder = @"C:\Users\ABSOLUTE PATH TO ExcelTemplate folder CryptDecrypt\ExcelTemplate";

            string template_fileName = @"decrypt.xlsx";
            var urlFile = Path.Combine(rootFolder, template_fileName);
            FileInfo files = new FileInfo(urlFile);
            var xlPackage = new ExcelPackage(files);
            ExcelWorksheet Sheet = xlPackage.Workbook.Worksheets["decrypt"];
            int index = 2;

            foreach (var r in list)
            {
                Sheet.Cells[index, 1].Value = r.id;
                Sheet.Cells[index, 2].Value = r.email_address;
                Sheet.Cells[index, 3].Value = r.migrated_password;
                Sheet.Cells[index, 4].Value = r.first_name;
                Sheet.Cells[index, 5].Value = r.last_name;
                Sheet.Cells[index, 6].Value = r.country_id;

                index = index + 1;
            }

            //To-Do
            var returnfilepath = Path.Combine(@"C:\Users..ABSOLUTE PATH to download folder\CryptDecrypt\downloads", filename);
            try
            {
                xlPackage.SaveAs(new FileInfo(returnfilepath));
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        static List<Student> GetDataTableFromDB(string cmdText)
        {
            //Data source details, oracle
            string con = "Data Source=(DESCRIPTION="
                   + "(ADDRESS=(PROTOCOL=TCP)(HOST=host_ip)(PORT=1540))"
                   + "(CONNECT_DATA=(SERVICE_NAME=Service_name)));"
                   + "User Id=user_id;Password=password;Connection Timeout=120;";

            List<Student> matchingPerson = new List<Student>();
            using (OracleConnection myConnection = new OracleConnection(con))
            {
                OracleCommand oCmd = new OracleCommand(cmdText, myConnection);
                myConnection.Open();
                using (OracleDataReader oReader = oCmd.ExecuteReader())
                {
                    while (oReader.Read())
                    {
                        Student Person = new Student();
                        Person.migrated_password = !string.IsNullOrEmpty(Convert.ToString(oReader["migrated_password"]).Trim()) ? Crypt.Decrypt(Convert.ToString(oReader["migrated_password"]).Trim()) : string.Empty;
                        Person.id = !string.IsNullOrEmpty(Convert.ToString(oReader["id"]).Trim()) ? Convert.ToString(oReader["id"]).Trim() : string.Empty;
                        Person.first_name = !string.IsNullOrEmpty(Convert.ToString(oReader["first_name"]).Trim()) ? Convert.ToString(oReader["first_name"]).Trim() : string.Empty;
                        Person.last_name = !string.IsNullOrEmpty(Convert.ToString(oReader["last_name"]).Trim()) ? Convert.ToString(oReader["last_name"]).Trim() : string.Empty;
                        Person.email_address = !string.IsNullOrEmpty(Convert.ToString(oReader["email_address"]).Trim()) ? Convert.ToString(oReader["email_address"]).Trim() : string.Empty;
                        Person.country_id = !string.IsNullOrEmpty(Convert.ToString(oReader["country_id"]).Trim()) ? Convert.ToString(oReader["country_id"]).Trim() : string.Empty;
                        matchingPerson.Add(Person);
                    }
                    myConnection.Close();
                }
            }
            return matchingPerson;
        }
    }

    internal class Student
    {
        public string migrated_password { get; internal set; }
        public string id { get; internal set; }
        public string first_name { get; internal set; }
        public string last_name { get; internal set; }
        public string email_address { get; internal set; }
        public string country_id { get; internal set; }
    }
}
