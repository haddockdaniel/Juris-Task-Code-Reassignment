using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Data;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Globalization;
using Gizmox.Controls;
using JDataEngine;
using JurisAuthenticator;
using JurisUtilityBase.Properties;
using System.Data.OleDb;

namespace JurisUtilityBase
{
    public partial class UtilityBaseMain : Form
    {
        #region Private  members

        private JurisUtility _jurisUtility;

        #endregion

        #region Public properties

        public string CompanyCode { get; set; }

        public string JurisDbName { get; set; }

        public string JBillsDbName { get; set; }

        public string fromTaskCode = "";

        public string toTaskCode = "";

        #endregion

        #region Constructor

        public UtilityBaseMain()
        {
            InitializeComponent();
            _jurisUtility = new JurisUtility();
        }

        #endregion

        #region Public methods

        public void LoadCompanies()
        {
            var companies = _jurisUtility.Companies.Cast<object>().Cast<Instance>().ToList();
//            listBoxCompanies.SelectedIndexChanged -= listBoxCompanies_SelectedIndexChanged;
            listBoxCompanies.ValueMember = "Code";
            listBoxCompanies.DisplayMember = "Key";
            listBoxCompanies.DataSource = companies;
//            listBoxCompanies.SelectedIndexChanged += listBoxCompanies_SelectedIndexChanged;
            var defaultCompany = companies.FirstOrDefault(c => c.Default == Instance.JurisDefaultCompany.jdcJuris);
            if (companies.Count > 0)
            {
                listBoxCompanies.SelectedItem = defaultCompany ?? companies[0];
            }
        }

        #endregion

        #region MainForm events

        private void Form1_Load(object sender, EventArgs e)
        {
        }

        private void listBoxCompanies_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (_jurisUtility.DbOpen)
            {
                _jurisUtility.CloseDatabase();
            }
            CompanyCode = "Company" + listBoxCompanies.SelectedValue;
            _jurisUtility.SetInstance(CompanyCode);
            JurisDbName = _jurisUtility.Company.DatabaseName;
            JBillsDbName = "JBills" + _jurisUtility.Company.Code;
            _jurisUtility.OpenDatabase();
            if (_jurisUtility.DbOpen)
            {
                ///GetFieldLengths();
            }

            string taskCode;
            cbFrom.ClearItems();
            string SQLTkpr = "select TaskCdCode + ' ' + TaskCdDesc as taskcode from TaskCode";
            DataSet myRSTkpr = _jurisUtility.RecordsetFromSQL(SQLTkpr);

            if (myRSTkpr.Tables[0].Rows.Count == 0)
                cbFrom.SelectedIndex = 0;
            else
            {
                foreach (DataTable table in myRSTkpr.Tables)
                {

                    foreach (DataRow dr in table.Rows)
                    {
                        taskCode = dr["taskcode"].ToString();
                        cbFrom.Items.Add(taskCode);
                    }
                }

            }

            string TkprIndex2;
            cbTo.ClearItems();
            string SQLTkpr2 = "select TaskCdCode + ' ' + TaskCdDesc as taskcode from TaskCode";
            DataSet myRSTkpr2 = _jurisUtility.RecordsetFromSQL(SQLTkpr2);


            if (myRSTkpr2.Tables[0].Rows.Count == 0)
                cbTo.SelectedIndex = 0;
            else
            {
                foreach (DataTable table in myRSTkpr2.Tables)
                {

                    foreach (DataRow dr in table.Rows)
                    {
                        TkprIndex2 = dr["taskcode"].ToString();
                        cbTo.Items.Add(TkprIndex2);
                    }
                }

            }




        }



        #endregion

        #region Private methods

        private void DoDaFix()
        {
            // Enter your SQL code here
            // To run a T-SQL statement with no results, int RecordsAffected = _jurisUtility.ExecuteNonQueryCommand(0, SQL);
            // To get an ADODB.Recordset, ADODB.Recordset myRS = _jurisUtility.RecordsetFromSQL(SQL);

            DialogResult dr = MessageBox.Show("This will change all Task code references from " + fromTaskCode + "\r\n" + "to " + toTaskCode + ". Are you sure?","Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (dr == System.Windows.Forms.DialogResult.Yes)
            {
                string SQL = "update ARFTaskAlloc set ARFTTaskCd ='" + toTaskCode + "' where ARFTTaskCd = '" + fromTaskCode + "'";
                _jurisUtility.ExecuteNonQueryCommand(0, SQL);
                UpdateStatus("Updating AR Fee Task and Billed Time...", 1, 8);

                SQL = "update BilledTime set BTTaskCd ='" + toTaskCode + "' where BTTaskCd = '" + fromTaskCode + "'";
                _jurisUtility.ExecuteNonQueryCommand(0, SQL);
                UpdateStatus("Updating Cash Rec Fee Alloc...", 2, 8);

                SQL = "update CRFeeAlloc set CRFTaskCd ='" + toTaskCode + "' where CRFTaskCd = '" + fromTaskCode + "'";
                _jurisUtility.ExecuteNonQueryCommand(0, SQL);
                UpdateStatus("Updating Fee Sum By Period...", 3, 8);

                SQL = "update FeeSumByPrd set FSPTaskCd ='" + toTaskCode + "' where FSPTaskCd = '" + fromTaskCode + "'";
                _jurisUtility.ExecuteNonQueryCommand(0, SQL);
                UpdateStatus("Updating Matter Fee Budget...", 4, 8);

                SQL = "update MatterFeeBudget set MFBTaskCode ='" + toTaskCode + "' where MFBTaskCode = '" + fromTaskCode + "'";
                _jurisUtility.ExecuteNonQueryCommand(0, SQL);
                UpdateStatus("Updating Time Batch Detail...", 5, 8);

                SQL = "update TimeBatchDetail set TBDTaskCd ='" + toTaskCode + "' where TBDTaskCd = '" + fromTaskCode + "'";
                _jurisUtility.ExecuteNonQueryCommand(0, SQL);
                UpdateStatus("Updating Time Entries...", 6, 8);

                SQL = "update TimeEntry set TaskCode ='" + toTaskCode + "' where TaskCode = '" + fromTaskCode + "'";
                _jurisUtility.ExecuteNonQueryCommand(0, SQL);
                UpdateStatus("Updating Unbilled Time...", 7, 8);

                SQL = "update UnbilledTime set UTTaskCd ='" + toTaskCode + "' where UTTaskCd = '" + fromTaskCode + "'";
                _jurisUtility.ExecuteNonQueryCommand(0, SQL);
                UpdateStatus("All tables updated.", 8, 8);

                //log tables?
                //xref?
                //fave and most recent?

                MessageBox.Show("The process is complete", "Finished", MessageBoxButtons.OK, MessageBoxIcon.None);
                toTaskCode = "";
                fromTaskCode = "";
                button1.Enabled = false;
            }
        }
        private bool VerifyFirmName()
        {
            //    Dim SQL     As String
            //    Dim rsDB    As ADODB.Recordset
            //
            //    SQL = "SELECT CASE WHEN SpTxtValue LIKE '%firm name%' THEN 'Y' ELSE 'N' END AS Firm FROM SysParam WHERE SpName = 'FirmName'"
            //    Cmd.CommandText = SQL
            //    Set rsDB = Cmd.Execute
            //
            //    If rsDB!Firm = "Y" Then
            return true;
            //    Else
            //        VerifyFirmName = False
            //    End If

        }

        private bool FieldExistsInRS(DataSet ds, string fieldName)
        {

            foreach (DataColumn column in ds.Tables[0].Columns)
            {
                if (column.ColumnName.Equals(fieldName, StringComparison.OrdinalIgnoreCase))
                    return true;
            }
            return false;
        }


        private static bool IsDate(String date)
        {
            try
            {
                DateTime dt = DateTime.Parse(date);
                return true;
            }
            catch
            {
                return false;
            }
        }

        private static bool IsNumeric(object Expression)
        {
            double retNum;

            bool isNum = Double.TryParse(Convert.ToString(Expression), System.Globalization.NumberStyles.Any, System.Globalization.NumberFormatInfo.InvariantInfo, out retNum);
            return isNum; 
        }

        private void WriteLog(string comment)
        {
            var sql =
                string.Format("Insert Into UtilityLog(ULTimeStamp,ULWkStaUser,ULComment) Values('{0}','{1}', '{2}')",
                    DateTime.Now, GetComputerAndUser(), comment);
            _jurisUtility.ExecuteNonQueryCommand(0, sql);
        }

        private string GetComputerAndUser()
        {
            var computerName = Environment.MachineName;
            var windowsIdentity = System.Security.Principal.WindowsIdentity.GetCurrent();
            var userName = (windowsIdentity != null) ? windowsIdentity.Name : "Unknown";
            return computerName + "/" + userName;
        }

        /// <summary>
        /// Update status bar (text to display and step number of total completed)
        /// </summary>
        /// <param name="status">status text to display</param>
        /// <param name="step">steps completed</param>
        /// <param name="steps">total steps to be done</param>
        private void UpdateStatus(string status, long step, long steps)
        {
            labelCurrentStatus.Text = status;

            if (steps == 0)
            {
                progressBar.Value = 0;
                labelPercentComplete.Text = string.Empty;
            }
            else
            {
                double pctLong = Math.Round(((double)step/steps)*100.0);
                int percentage = (int)Math.Round(pctLong, 0);
                if ((percentage < 0) || (percentage > 100))
                {
                    progressBar.Value = 0;
                    labelPercentComplete.Text = string.Empty;
                }
                else
                {
                    progressBar.Value = percentage;
                    labelPercentComplete.Text = string.Format("{0} percent complete", percentage);
                }
            }
        }

        private void DeleteLog()
        {
            string AppDir = Path.GetDirectoryName(Application.ExecutablePath);
            string filePathName = Path.Combine(AppDir, "VoucherImportLog.txt");
            if (File.Exists(filePathName + ".ark5"))
            {
                File.Delete(filePathName + ".ark5");
            }
            if (File.Exists(filePathName + ".ark4"))
            {
                File.Copy(filePathName + ".ark4", filePathName + ".ark5");
                File.Delete(filePathName + ".ark4");
            }
            if (File.Exists(filePathName + ".ark3"))
            {
                File.Copy(filePathName + ".ark3", filePathName + ".ark4");
                File.Delete(filePathName + ".ark3");
            }
            if (File.Exists(filePathName + ".ark2"))
            {
                File.Copy(filePathName + ".ark2", filePathName + ".ark3");
                File.Delete(filePathName + ".ark2");
            }
            if (File.Exists(filePathName + ".ark1"))
            {
                File.Copy(filePathName + ".ark1", filePathName + ".ark2");
                File.Delete(filePathName + ".ark1");
            }
            if (File.Exists(filePathName ))
            {
                File.Copy(filePathName, filePathName + ".ark1");
                File.Delete(filePathName);
            }

        }

            

        private void LogFile(string LogLine)
        {
            string AppDir = Path.GetDirectoryName(Application.ExecutablePath);
            string filePathName = Path.Combine(AppDir, "VoucherImportLog.txt");
            using (StreamWriter sw = File.AppendText(filePathName))
            {
                sw.WriteLine(LogLine);
            }	
        }
        #endregion

        private void button1_Click(object sender, EventArgs e)
        {
            DoDaFix();
        }

        private void buttonReport_Click(object sender, EventArgs e)
        {

            System.Environment.Exit(0);
          
        }

        private string getReportSQL()
        {
            string reportSQL = "";
            //if matter and billing timekeeper
            if (true)
                reportSQL = "select Clicode, Clireportingname, Matcode, Matreportingname,empinitials as CurrentBillingTimekeeper, 'DEF' as NewBillingTimekeeper" +
                        " from matter" +
                        " inner join client on matclinbr=clisysnbr" +
                        " inner join billto on matbillto=billtosysnbr" +
                        " inner join employee on empsysnbr=billtobillingatty" +
                        " where empinitials<>'ABC'";


            //if matter and originating timekeeper
            else if (false)
                reportSQL = "select Clicode, Clireportingname, Matcode, Matreportingname,empinitials as CurrentOriginatingTimekeeper, 'DEF' as NewOriginatingTimekeeper" +
                    " from matter" +
                    " inner join client on matclinbr=clisysnbr" +
                    " inner join matorigatty on matsysnbr=morigmat" +
                    " inner join employee on empsysnbr=morigatty" +
                    " where empinitials<>'ABC'";


            return reportSQL;
        }

        private void cbFrom_SelectedIndexChanged(object sender, EventArgs e)
        {
            fromTaskCode = cbFrom.Text;
            fromTaskCode = fromTaskCode.Split(' ')[0];
            if (!String.IsNullOrEmpty(toTaskCode))
                button1.Enabled = true;
        }

        private void cbTo_SelectedIndexChanged(object sender, EventArgs e)
        {
            toTaskCode = cbTo.Text;
            toTaskCode = toTaskCode.Split(' ')[0];
            if (!String.IsNullOrEmpty(fromTaskCode))
                button1.Enabled = true;
        }


    }
}
