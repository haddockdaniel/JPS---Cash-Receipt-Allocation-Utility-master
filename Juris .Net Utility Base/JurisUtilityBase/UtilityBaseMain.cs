using System;
using System.Collections.Generic;
using System.Linq;
using System.Data;
using System.IO;
using System.Windows.Forms;
using JurisAuthenticator;
using System.ComponentModel;
using System.Threading;
using System.Reflection;
using System.Data.SqlClient; 

namespace JurisUtilityBase
{
    public partial class UtilityBaseMain : Form
    {
        #region Private  members

        private JurisUtility _jurisUtility;

        #endregion

        #region Public properties
        //152557.82
        public string CompanyCode { get; set; }

        public string JurisDbName { get; set; }

        public string JBillsDbName { get; set; }

        public int FldClient { get; set; }

        public int FldMatter { get; set; }

        List<Check> checks = new List<Check>();

        List<DataError> errors = new List<DataError>();
        bool selectedSpreadsheet = false;

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
            string sql = "select left(empinitials + '         ',8) + empname as UName from employee where empvalidasuser='Y' order by case when empinitials='SMGR' then 1 else rank() over (order by empname) + 1 end ";
            DataSet dt = _jurisUtility.RecordsetFromSQL(sql);



            DataTable dtFS = dt.Tables[0];

            if (dtFS.Rows.Count == 0)
                cbUser.SelectedIndex = 0;
            else
            {
                string FSIndex = "";
                foreach (DataRow dr in dtFS.Rows)
                {
                    FSIndex = dr["UName"].ToString();
                    cbUser.Items.Add(FSIndex);
                }
            }
            cbUser.SelectedIndex = 0;
            getSettings();
        }



        #endregion

        #region Private methods

        private void DoDaFix()
        {
            if (selectedSpreadsheet)
            {
                button1.Enabled = false;
                buttonReport.Enabled = false;
                button1.ForeColor = System.Drawing.Color.Navy;
                Cursor.Current = Cursors.WaitCursor;
                toolStripStatusLabel.Text = "Creating Cash Receipt Records....";
                statusStrip.Refresh();
                UpdateStatus("Creating Cash Receipt Records....", 1, 10);
                Application.DoEvents();

                string BatRecAmount = "";
                string SQLC = "select max(case when spname='CurAcctPrdYear' then cast(spnbrvalue as varchar(4)) else '' end) as PrdYear, max(Case when spname = 'CurAcctPrdNbr' then case " +
                    " when spnbrvalue<9 then '0' + cast(spnbrvalue as varchar(1)) else cast(spnbrvalue as varchar(2)) end  else '' end) as PrdNbr," +
                    "max(case when spname='CfgMiscOpts' then substring(sptxtvalue,14,1) else 0 end) as DOrder from sysparam";
                DataSet myRSSysParm = _jurisUtility.RecordsetFromSQL(SQLC);

                DataTable dtSP = myRSSysParm.Tables[0];

                if (dtSP.Rows.Count == 0)
                { MessageBox.Show("Incorrect SysParams"); }
                else
                {
                    foreach (DataRow dr in dtSP.Rows)
                    {
                        PYear = dr["PrdYear"].ToString();
                        PNbr = dr["PrdNbr"].ToString();
                        DOrder = dr["DOrder"].ToString();

                    }
                }

                //##CombineChecks (Deposit_Date, Check_Number, Check_Amount, Check_Date, [Payor], " +
                //" Client_Code, Matter_Code, Bank_Code, GL_Account, [Reference], " +
                /////" Invoice_Date, Invoice_Number, Bill_Total, Bill_Balance, AR_Allocation, Fees, Cash_Exps, NonCash_Exps, Taxes1, " +
                //" Taxes2, Taxes3, Surcharge, Interest, PPD_Allocation, Trust_Allocation, Other_Allocation, TTYpe, batchNo

                CreateBatch();

                Cursor.Current = Cursors.Default;
                toolStripStatusLabel.Text = "Utility Completed.";
                statusStrip.Refresh();
                UpdateStatus("Utility Completed.", 1, 1);
                Application.DoEvents();

                string cmt = Application.ProductName.ToString();
                WriteLog(cmt);

                MessageBox.Show("Cash Receipt Allocation Completed.");
                selectedSpreadsheet = false;
                checks.Clear();
                errors.Clear();
                button1.Enabled = false;
                buttonReport.Enabled = true;
                button1.ForeColor = System.Drawing.Color.Navy;
            }
            else
                MessageBox.Show("Please select a spreadsheet file first.");
        }

        private void CreateBatch()
        {
            Cursor.Current = Cursors.WaitCursor;
            toolStripStatusLabel.Text = "Creating Cash Receipt Batch...";
            statusStrip.Refresh();
            UpdateStatus("Creating Cash Receipt Batch...", 1, 10);
            Application.DoEvents();

            string singleBatch = "";

            //get distinct dep date and check total per date and create batch then loop through then update table with batch numbers
            string sql = "IF OBJECT_ID('tempdb.dbo.##DatesForBatches', 'U') IS NOT NULL DROP TABLE [dbo].[##DatesForBatches]";
            _jurisUtility.ExecuteNonQuery(0, sql);

            sql = "create table ##DatesForBatches (DepDate datetime, CheckNumber varchar(30), CheckAmt money)";
            _jurisUtility.ExecuteNonQuery(0, sql);

            sql = "insert into ##DatesForBatches (DepDate, CheckNumber, CheckAmt) select distinct Deposit_Date, Check_Number, Check_Amount from ##CombineChecks";
            _jurisUtility.ExecuteNonQuery(0, sql);

            //see how many batches are needed for progress bar
            sql = "";

            sql = "select DepDate, sum(CheckAmt) as CheckAmt from ##DatesForBatches group by DepDate";
            DataSet batches = _jurisUtility.RecordsetFromSQL(sql);

            foreach (DataRow dr in batches.Tables[0].Rows)
            {
                string BatRecAmount = dr[1].ToString();
                DateTime ddd = Convert.ToDateTime(dr[0].ToString());
                string BatDepDate = ddd.ToShortDateString();
                string MYFolder = PYear + "-" + PNbr;


                string STest = "select crbbatchnbr from CashReceiptsBatch where crbstatus='U' and crbenteredby=1 and crbcomment like 'CR Alloc Tool-' + convert(varchar,'" + BatDepDate + "', 101) + '%'  ";
                DataSet DT = _jurisUtility.RecordsetFromSQL(STest);
                DataTable d1 = DT.Tables[0];

                if (d1.Rows.Count == 0)
                {
                    string SQL = "Insert into CashReceiptsBatch(crbbatchnbr, crbcomment, crbstatus, crbreccount, crbenteredby,crbdateentered, crbpostedby, crbdateposted, crbbatchtotal)" +
                         " Values( (select spnbrvalue from sysparam where spname='LastBatchCash') + 1 ,left('CR Alloc Tool-' + convert(varchar,'" + BatDepDate + "', 101) + ' '  + cast((select spnbrvalue from sysparam where spname='LastBatchCash') + 1 as varchar(20)), 50)," +
                         "'U' , 1 ," + EmpSys.ToString() + ",convert(varchar(10),getdate(),101) , " + EmpSys.ToString() + " , convert(varchar(10),getdate(),101) , cast('" + BatRecAmount + "' as money) )";
                    _jurisUtility.ExecuteNonQueryCommand(0, SQL);

                    SQL = "Update sysparam set spnbrvalue=spnbrvalue + 1 where spname='LastBatchCash'";
                    _jurisUtility.ExecuteNonQueryCommand(0, SQL);

                    SQL = "select max(case when spname='CurAcctPrdYear' then cast(spnbrvalue as varchar(4)) else '' end) as PrdYear, " +
                       "max(Case when spname='CurAcctPrdNbr' then case when spnbrvalue<9 then '0' + cast(spnbrvalue as varchar(1)) else cast(spnbrvalue as varchar(2)) end  else '' end) as PrdNbr, " +
                       "max(case when spname='LastSysNbrDocTree' then spnbrvalue else 0 end) as DTree,max(case when spname='CfgMiscOpts' then substring(sptxtvalue,14,1) else 0 end) as DOrder from sysparam";
                    DataSet myRSSysParm = _jurisUtility.RecordsetFromSQL(SQL);

                    DataTable dtSP = myRSSysParm.Tables[0];

                    if (dtSP.Rows.Count == 0)
                    { MessageBox.Show("Incorrect SysParams"); }
                    else
                    {
                        foreach (DataRow dr3 in dtSP.Rows)
                        {
                            string LastSys = dr3["DTree"].ToString();
                            DOrder = dr3["DOrder"].ToString();
                            if (DOrder == "2")
                            {
                                string SPSql = "Select dtdocid from documenttree where dtparentid=35 and dtdocclass='5300' and dttitle='" + MYFolder + "'";
                                DataSet spMY = _jurisUtility.RecordsetFromSQL(SPSql);
                                DataTable dtMY = spMY.Tables[0];
                                if (dtMY.Rows.Count == 0)
                                {
                                    string s2 = "Insert into documenttree(dtdocid, dtsystemcreated, dtdocclass, dtdoctype, dtparentid, dttitle) " +
                                          "select max(dtdocid)  + 1, 'Y', 5300,'F', 35,'" + MYFolder + "' from documenttree ";
                                    _jurisUtility.ExecuteNonQueryCommand(0, s2);
                                    s2 = "Update sysparam set spnbrvalue=(select max(dtdocid) from documenttree) where spname='LastSysNbrDocTree'";
                                    _jurisUtility.ExecuteNonQueryCommand(0, s2);

                                    s2 = "Insert into documenttree(dtdocid, dtsystemcreated, dtdocclass, dtdoctype, dtparentid, dttitle) " +
                                        "select (select max(dtdocid) from documenttree) + 1, 'Y', 5300,'F', dtdocid,'" + EmpInt + "'" +
                                        " from documenttree where dtparentid=35 and dttitle='" + MYFolder + "'";
                                    _jurisUtility.ExecuteNonQueryCommand(0, s2);

                                    s2 = "Update sysparam set spnbrvalue=(select max(dtdocid) from documenttree) where spname='LastSysNbrDocTree'";
                                    _jurisUtility.ExecuteNonQueryCommand(0, s2);

                                    s2 = "Insert into documenttree(dtdocid, dtsystemcreated, dtdocclass, dtdoctype, dtparentid, dttitle, dtkeyL) " +
                                        "select (select max(dtdocid) from documenttree) + 1, 'Y', 5300,'R', " +
                                        " (Select dtdocid from documenttree where dtparentid=(Select dtdocid from documenttree where dtparentid=35  and dttitle='" + MYFolder + "') and dttitle='" + EmpInt + "')," +
                                            "left('CR Alloc Tool-' + convert(varchar,'" + BatDepDate + "', 101)+ '-' + cast((select spnbrvalue from sysparam where spname='LastBatchCash') as varchar(20)), 30), " +
                                            "cast((select spnbrvalue from sysparam where spname='LastBatchCash') as int)";
                                    _jurisUtility.ExecuteNonQueryCommand(0, s2);


                                    s2 = "Update sysparam set spnbrvalue=(select max(dtdocid) from documenttree) where spname='LastSysNbrDocTree'";
                                    _jurisUtility.ExecuteNonQueryCommand(0, s2);
                                }
                                else
                                {
                                    string SMGRSql = "Select dtdocid from documenttree where dtparentid=(Select dtdocid from documenttree where dtparentid=35 and dttitle='" + MYFolder + "') and dttitle='" + EmpInt + "'";
                                    DataSet spSMGR = _jurisUtility.RecordsetFromSQL(SMGRSql);
                                    DataTable dtSMGR = spSMGR.Tables[0];
                                    if (dtSMGR.Rows.Count == 0)
                                    {
                                        string s2 = "Insert into documenttree(dtdocid, dtsystemcreated, dtdocclass, dtdoctype, dtparentid, dttitle) " +
                                       "select (select max(dtdocid) from documenttree) + 1, 'Y', 5300,'F', dtdocid,'" + EmpInt + "'" +
                                       " from documenttree where dtparentid=35 and dttitle='" + MYFolder + "'";
                                        _jurisUtility.ExecuteNonQueryCommand(0, s2);

                                        s2 = "Update sysparam set spnbrvalue=(select max(dtdocid) from documenttree) where spname='LastSysNbrDocTree'";
                                        _jurisUtility.ExecuteNonQueryCommand(0, s2);

                                        s2 = "Insert into documenttree(dtdocid, dtsystemcreated, dtdocclass, dtdoctype, dtparentid, dttitle, dtkeyL) " +
                                            "select (select max(dtdocid) from documenttree) + 1, 'Y', 5300,'R', " +
                                            " (Select dtdocid from documenttree where dtparentid=(Select dtdocid from documenttree where dtparentid=35 and dttitle='" + MYFolder + "')  and dttitle='" + EmpInt + "')," +
                                            "left('CR Alloc Tool-' + convert(varchar,'" + BatDepDate + "', 101)+ '-' + cast((select spnbrvalue from sysparam where spname='LastBatchCash') as varchar(20)), 30), " +
                                            "cast((select spnbrvalue from sysparam where spname='LastBatchCash') as int)";
                                        _jurisUtility.ExecuteNonQueryCommand(0, s2);


                                        s2 = "Update sysparam set spnbrvalue=(select max(dtdocid) from documenttree) where spname='LastSysNbrDocTree'";
                                        _jurisUtility.ExecuteNonQueryCommand(0, s2);
                                    }
                                    else
                                    {
                                        string s2 = "Insert into documenttree(dtdocid, dtsystemcreated, dtdocclass, dtdoctype, dtparentid, dttitle, dtkeyL) " +
                                            "select (select max(dtdocid) from documenttree) + 1, 'Y', 5300,'R', " +
                                            " (Select dtdocid from documenttree where dtparentid=(Select dtdocid from documenttree where dtparentid=35 and dttitle='" + MYFolder + "')  and dttitle='" + EmpInt + "')," +
                                            "left('CR Alloc Tool-' + convert(varchar,'" + BatDepDate + "', 101)+ '-' + cast((select spnbrvalue from sysparam where spname='LastBatchCash') as varchar(20)), 30), " +
                                            "cast((select spnbrvalue from sysparam where spname='LastBatchCash') as int)";
                                        _jurisUtility.ExecuteNonQueryCommand(0, s2);


                                        s2 = "Update sysparam set spnbrvalue=(select max(dtdocid) from documenttree) where spname='LastSysNbrDocTree'";
                                        _jurisUtility.ExecuteNonQueryCommand(0, s2);
                                    }
                                }
                            }
                            else
                            {
                                string SPSql = "Select dtdocid from documenttree where dtparentid=35 and dtdocclass='5300' and dttitle='" + EmpInt + "'";
                                DataSet spMY = _jurisUtility.RecordsetFromSQL(SPSql);
                                DataTable dtMY = spMY.Tables[0];
                                if (dtMY.Rows.Count == 0)
                                {
                                    string s2 = "Insert into documenttree(dtdocid, dtsystemcreated, dtdocclass, dtdoctype, dtparentid, dttitle) " +
                                          "select max(dtdocid)  + 1, 'Y', 5300,'F', 35,'" + EmpInt + "' from documenttree ";
                                    _jurisUtility.ExecuteNonQueryCommand(0, s2);
                                    s2 = "Update sysparam set spnbrvalue=(select max(dtdocid) from documenttree) where spname='LastSysNbrDocTree'";
                                    _jurisUtility.ExecuteNonQueryCommand(0, s2);

                                    s2 = "Insert into documenttree(dtdocid, dtsystemcreated, dtdocclass, dtdoctype, dtparentid, dttitle) " +
                                        "select (select max(dtdocid) from documenttree) + 1, 'Y', 5300,'F', dtdocid,'" + MYFolder + "'" +
                                        " from documenttree where dtparentid=35 and dttitle='" + EmpInt + "'";
                                    _jurisUtility.ExecuteNonQueryCommand(0, s2);

                                    s2 = "Update sysparam set spnbrvalue=(select max(dtdocid) from documenttree) where spname='LastSysNbrDocTree'";
                                    _jurisUtility.ExecuteNonQueryCommand(0, s2);

                                    s2 = "Insert into documenttree(dtdocid, dtsystemcreated, dtdocclass, dtdoctype, dtparentid, dttitle, dtkeyL) " +
                                        "select (select max(dtdocid) from documenttree) + 1, 'Y', 5300,'R', " +
                                        " (Select dtdocid from documenttree where dtparentid=(Select dtdocid from documenttree where dtparentid=35  and dttitle='" + EmpInt + "') and dttitle='" + MYFolder + "')," +
                                            "left('CR Alloc Tool-' + convert(varchar,'" + BatDepDate + "', 101)+ '-' + cast((select spnbrvalue from sysparam where spname='LastBatchCash') as varchar(20)), 30), " +
                                            "cast((select spnbrvalue from sysparam where spname='LastBatchCash') as int)";
                                    _jurisUtility.ExecuteNonQueryCommand(0, s2);


                                    s2 = "Update sysparam set spnbrvalue=(select max(dtdocid) from documenttree) where spname='LastSysNbrDocTree'";
                                    _jurisUtility.ExecuteNonQueryCommand(0, s2);
                                }
                                else
                                {
                                    string SMGRSql = "Select dtdocid from documenttree where dtparentid=(Select dtdocid from documenttree where dtparentid=35  and dttitle='" + EmpInt + "') and dttitle='" + MYFolder + "'";
                                    DataSet spSMGR = _jurisUtility.RecordsetFromSQL(SMGRSql);
                                    DataTable dtSMGR = spSMGR.Tables[0];
                                    if (dtSMGR.Rows.Count == 0)
                                    {
                                        string s2 = "Insert into documenttree(dtdocid, dtsystemcreated, dtdocclass, dtdoctype, dtparentid, dttitle) " +
                                       "select (select max(dtdocid) from documenttree) + 1, 'Y', 5300,'F', dtdocid,'" + MYFolder + "'" +
                                       " from documenttree where dtparentid=35 and dttitle='" + EmpInt + "'";
                                        _jurisUtility.ExecuteNonQueryCommand(0, s2);

                                        s2 = "Update sysparam set spnbrvalue=(select max(dtdocid) from documenttree) where spname='LastSysNbrDocTree'";
                                        _jurisUtility.ExecuteNonQueryCommand(0, s2);

                                        s2 = "Insert into documenttree(dtdocid, dtsystemcreated, dtdocclass, dtdoctype, dtparentid, dttitle, dtkeyL) " +
                                            "select (select max(dtdocid) from documenttree) + 1, 'Y', 5300,'R', " +
                                            " (Select dtdocid from documenttree where dtparentid=(Select dtdocid from documenttree where dtparentid=35   and dttitle='" + EmpInt + "')and dttitle='" + MYFolder + "')," +
                                            "left('CR Alloc Tool-' + convert(varchar,'" + BatDepDate + "', 101)+ '-' + cast((select spnbrvalue from sysparam where spname='LastBatchCash') as varchar(20)), 30), " +
                                            "cast((select spnbrvalue from sysparam where spname='LastBatchCash') as int)";
                                        _jurisUtility.ExecuteNonQueryCommand(0, s2);


                                        s2 = "Update sysparam set spnbrvalue=(select max(dtdocid) from documenttree) where spname='LastSysNbrDocTree'";
                                        _jurisUtility.ExecuteNonQueryCommand(0, s2);
                                    }
                                    else
                                    {
                                        string s2 = "Insert into documenttree(dtdocid, dtsystemcreated, dtdocclass, dtdoctype, dtparentid, dttitle, dtkeyL) " +
                                            "select (select max(dtdocid) from documenttree) + 1, 'Y', 5300,'R', " +
                                            " (Select dtdocid from documenttree where dtparentid=(Select dtdocid from documenttree where dtparentid=35 and dttitle='" + EmpInt + "') and dttitle='" + MYFolder + "') ," +
                                            "left('CR Alloc Tool-' + convert(varchar,'" + BatDepDate + "', 101)+ '-' + cast((select spnbrvalue from sysparam where spname='LastBatchCash') as varchar(20)), 30), " +
                                            "cast((select spnbrvalue from sysparam where spname='LastBatchCash') as int)";
                                        _jurisUtility.ExecuteNonQueryCommand(0, s2);


                                        s2 = "Update sysparam set spnbrvalue=(select max(dtdocid) from documenttree) where spname='LastSysNbrDocTree'";
                                        _jurisUtility.ExecuteNonQueryCommand(0, s2);
                                    }
                                }

                            }
                        }
                    }
                    string sqlB = "select crbbatchnbr from cashreceiptsbatch where crbcomment= left('CR Alloc Tool-' + convert(varchar,'" + BatDepDate + "',101) + ' '  + cast((select spnbrvalue from sysparam where spname='LastBatchCash') as varchar(20)), 50)  and crbstatus='U' and crbreccount=1 " +
                        " and convert(varchar(10),crbdateentered,101) =convert(varchar(10),getdate(),101)  and crbbatchtotal= cast('" + BatRecAmount + "' as money) ";
                    DataSet spBatch = _jurisUtility.RecordsetFromSQL(sqlB);
                    DataTable dtB = spBatch.Tables[0];
                    if (dtB.Rows.Count == 0)
                    { MessageBox.Show("Error Creating Cash Receipt Batch"); }
                    else
                    {
                        foreach (DataRow dr2 in dtB.Rows)
                        {
                            singleBatch = dr2["crbbatchnbr"].ToString();
                        }

                    }
                }

                else
                {
                    foreach (DataRow dr1 in d1.Rows)
                    {
                        singleBatch = dr1["crbbatchnbr"].ToString();
                        string s2 = "Update CashReceiptsBatch set crbreccount=crbreccount + 1 where crbbatchnbr=" + singleBatch;
                        _jurisUtility.ExecuteNonQueryCommand(0, s2);
                    }

                }

                sql = "update ##CombineChecks set batchNo = " + singleBatch + " where Deposit_Date = '" + BatDepDate + "'";
                _jurisUtility.ExecuteNonQueryCommand(0, sql);
                CreateBatchRecord(singleBatch, BatDepDate);
                CreateBatchPPD(singleBatch);
                CreateBatchTrust(singleBatch);
                CreateBatchOther(singleBatch);
                CreateBatchAR(singleBatch);
            }
        }

        private void CreateBatchRecord(string batchNo, string depDate)
        {
            Cursor.Current = Cursors.WaitCursor;
            toolStripStatusLabel.Text = "Creating Cash Receipt Batch Record for Batch " + batchNo + "...";
            statusStrip.Refresh();
            UpdateStatus("Creating Cash Receipt Batch Record for Batch " + batchNo + "...", 2, 10);
            Application.DoEvents();

            string sql = "select Check_Number, max(Check_Amount) as Check_Amount, max(Check_Date) as Check_Date, max([Payor]) as [Payor], " +
                " sum(AR_Allocation) as AR_Allocation, sum(PPD_Allocation) as PPD_Allocation, sum(Trust_Allocation) as Trust_Allocation, " +
                " sum(Other_Allocation) as NonCli from ##CombineChecks where batchNo = " + batchNo + " and Deposit_Date = '" + depDate + "' group by Check_Number ";

            DataSet records = _jurisUtility.RecordsetFromSQL(sql);

            string SQL = "";

            foreach (DataRow rrow in records.Tables[0].Rows)
            {
                     ///money or decimal?
              SQL = "Insert into cashreceipt(crbatch, crrecnbr,crposted, crdate, crprdyear, crprdnbr, crchecknbr, crcheckdate, crcheckamt, crpayor, crarcsh, crppdcsh, crtrustcsh, crnonclicsh)" +
                    "values (" + batchNo + ", case when(select max(crrecnbr) from CashReceipt where crbatch = " + batchNo + ") is null then 1 else ((select max(crrecnbr) from CashReceipt where crbatch = " + batchNo + ") +1) end " +
                    " ,'N',convert(datetime,'" + depDate + "',101) ," + PYear + "," + PNbr + ",'" + rrow[0].ToString() + "','" + rrow[2].ToString() + "',cast('" + rrow[1].ToString() + "' as money),'" + rrow[3].ToString() +
                  "',cast('" + rrow[4].ToString() + "' as money),cast('" + rrow[5].ToString() + "' as money),cast('" + rrow[6].ToString() + "' as money),cast('" + rrow[7].ToString() + "' as money) )";
                _jurisUtility.ExecuteNonQueryCommand(0, SQL);
            }

            //Update the batch record to have the right number of records
            SQL = "  update cb  set cb.CRBRecCount = cr.mmx " +
                  " from CashReceiptsBatch cb " +
                  " inner join (select CRBatch, max(CRRecNbr) as mmx  from CashReceipt where CRBatch = " + batchNo + " group by crbatch) cr on cr.CRBatch = cb.CRBBatchNbr";
            _jurisUtility.ExecuteNonQueryCommand(0, SQL);
        }

        private void CreateBatchAR(string batchNo)
        {
            Cursor.Current = Cursors.WaitCursor;
            toolStripStatusLabel.Text = "Creating Cash Receipt Batch AR Record for Batch " + batchNo + "...";
            statusStrip.Refresh();
            UpdateStatus("Creating Cash Receipt  Batch AR...", 3, 10);
            Application.DoEvents();

            string SQL = "Insert into CRARAlloc(crabatch, crarecnbr, cramatter, crabillnbr, cradate, crachecknbr, cracheckdate, crapayor, crafeeamt, cracshexpamt, crancshexpamt, CRASurchgAmt ,CRATax1Amt,CRATax2Amt,CRATax3Amt,CRAInterestAmt, craprepostfee, craprepostcshexp, craprepostncshexp, CRAPrePostSurchg ,CRAPrePostTax1 ,CRAPrePostTax2 ,CRAPrePostTax3,CRAPrePostInterest, crabank) " +
             " Select  batchNo,  crbreccount,  matsys, Invoice_Number,convert(datetime,Deposit_Date,101) ,Check_Number,Check_Date,[Payor],fees,  " +
             " cshexp, ncshexp, surcharge, tax1, tax2, tax3, interest, feeAR, cshAR, ncshAR, surAR, tax1AR, tax2AR, tax3AR, intAR, ofcbankcode  " +
             " from (select matsys, batchNo, Invoice_Number, Deposit_Date, Check_Number, Check_Date, [Payor], sum(Fees) as fees, sum(Cash_Exps) as cshexp, sum(NonCash_Exps) as NCshExp, sum(Taxes1) as tax1, sum(Taxes2) as tax2,  " +
             " sum(Taxes3) as tax3, sum(Surcharge) as surcharge, sum(Interest) as interest " +
            " from ##CombineChecks where ttype = 1 and batchno = " + batchNo + " and AR_Allocation is not null and AR_Allocation <> 0 group by matsys, batchNo,Invoice_Number, Deposit_Date, Check_Number, Check_Date, [Payor]) AR  " +
            " inner join matter on matsysnbr=matsys inner join officecode on matofficecode=ofcofficecode  " +
             " Inner join (select armmatter, armbillnbr, sum(armfeebld - armfeercvd + armfeeadj) as FeeAR, sum(armcshexpbld - armcshexprcvd + armcshexpadj) as CshAR,sum(armncshexpbld - armncshexprcvd + armncshexpadj) as ncshAR, sum(ARMSurchgBld-ARMSurchgRcvd+ARMSurchgAdj) as surAR,  " +
             " sum([ARMTax1Bld]-[ARMTax1Rcvd]+[ARMTax1Adj]) as tax1AR, sum([ARMTax2Bld]-[ARMTax2Rcvd]+[ARMTax2Adj]) as tax2AR, sum([ARMTax3Bld]-[ARMTax3Rcvd]+[ARMTax3Adj]) as tax3AR, sum([ARMIntBld]-[ARMIntRcvd]+[ARMIntAdj]) as intAR  " +
             " from armatalloc group by armmatter, armbillnbr) ARM on matsys=armmatter and Invoice_Number=armbillnbr, cashreceiptsbatch where crbbatchnbr= "+ batchNo;


            //SQL = "Insert into CRARAlloc(crabatch, crarecnbr, cramatter, crabillnbr, cradate, crachecknbr, cracheckdate, crapayor, crafeeamt, cracshexpamt, crancshexpamt, CRASurchgAmt ,CRATax1Amt,CRATax2Amt,CRATax3Amt,CRAInterestAmt, craprepostfee, craprepostcshexp, craprepostncshexp, CRAPrePostSurchg ,CRAPrePostTax1 ,CRAPrePostTax2 ,CRAPrePostTax3,CRAPrePostInterest, crabank) " +
           // " Select  crbbatchnbr, crbreccount, matter, billnbr,convert(datetime,'" + depDate + "',101) ,'" + CkNbr + "','" + CkDate + "','" + Payor + "',fees, cshexp, ncshexp, surcharge, tax`1, tax2, tax3, interest, feeAR, cshAR, ncshAR, surAR, tax1AR, tax2AR, tax3AR, intAR, ofcbankcode " +
           // " from (select matter, billnbr, sum(case when itype='Fee' then allocamt else 0 end) as fees, sum(case when ITYpe='Cost' and exptype='C' then allocamt else 0 end) as cshexp, sum(case when ITYpe='Cost' and exptype='N' then allocamt else 0 end) as NCshExp " +
           // " from #ARAlloc group by matter, billnbr) AR " +
          //  "inner join matter on matsysnbr=matter inner join officecode on matofficecode=ofcofficecode " +
          ////  " Inner join (select armmatter, armbillnbr, sum(armfeebld - armfeercvd + armfeeadj) as FeeAR, sum(armcshexpbld - armcshexprcvd + armcshexpadj) as CshAR,sum(armncshexpbld - armncshexprcvd + armncshexpadj) as ncshAR, sum([ARMSurchgBld]-[ARMSurchgRcvd]+[ARMSurchgAdj) as surAR, " +
         //   " sum([ARMTax1Bld]-[ARMTax1Rcvd]+[ARMTax1Adj]) as tax1AR, sum([ARMTax2Bld]-[ARMTax2Rcvd]+[ARMTax2Adj]) as tax2AR, sum([ARMTax3Bld]-[ARMTax3Rcvd]+[ARMTax3Adj]) as tax2AR, sum([ARMIntBld]-[ARMIntRcvd]+[ARMIntAdj]) as intAR " + 
         //   " from armatalloc group by armmatter, armbillnbr) ARM on matter=armmatter and billnbr=armbillnbr, cashreceiptsbatch where crbbatchnbr=" + batchNo;
            _jurisUtility.ExecuteNonQueryCommand(0, SQL);                           


            SQL = "Update ARMatalloc set armpendfee= crafeeamt, armpendcshexp=cracshexpamt, armpendncshexp=crancshexpamt,  " +
                " ARMPendSurchg = CRASurchgAmt, ARMPendTax1 = CRATax1Amt, ARMPendTax2 = CRATax2Amt, ARMPendTax3 = CRATax3Amt, ARMPendInt = CRAInterestAmt" +
                " from (select distinct cramatter, crabillnbr, crafeeamt, cracshexpamt, crancshexpamt, CRASurchgAmt, CRATax1Amt, CRATax2Amt, CRATax3Amt, CRAInterestAmt from craralloc " +
                " inner join ##CombineChecks on matsys=cramatter and Invoice_Number=crabillnbr and ttype = 1 and AR_Allocation is not null and AR_Allocation <> 0 and batchNo = " + batchNo + ")CR where cramatter=armmatter and crabillnbr=armbillnbr";

            _jurisUtility.ExecuteNonQueryCommand(0, SQL);

            string sql = "IF  NOT EXISTS (SELECT * FROM sys.objects " +
            " WHERE object_id = OBJECT_ID(N'[dbo].[#ARAlloc]') AND type in (N'U')) " +
            " BEGIN " +
            " create table [dbo].#ARAlloc (matter int, billnbr int, tkpr int, taskcd varchar(6), actcd varchar(6), expcd varchar(6), exptype varchar(6), allocAmt decimal(15,2) )" +
            " END";

            _jurisUtility.ExecuteSqlCommand(0, sql);

            CreateBatchFees(batchNo);
            CreateBatchExps(batchNo);

        }

        private void CreateBatchPPD(string batchNo)
        {
            Cursor.Current = Cursors.WaitCursor;
            toolStripStatusLabel.Text = "Creating Cash Receipt Batch PPD Record for Batch " + batchNo + "...";
            statusStrip.Refresh();
            UpdateStatus("Creating Cash Receipt  Batch PPD...", 4, 10);
            Application.DoEvents();
            string SQL = "";

            SQL = "select matsys, PPD_Allocation from ##CombineChecks where PPD_Allocation is not null and PPD_Allocation <> 0 and batchNo = " + batchNo;

            DataSet ppd = _jurisUtility.RecordsetFromSQL(SQL);

            foreach (DataRow alloc in ppd.Tables[0].Rows)
            {
                SQL = "Insert into CRPPDAlloc(crpbatch, crprecnbr, crpmatter, crpamount) " +
                        "values (" + batchNo + ", case when(select max(CRPRecNbr) from CRPPDAlloc where CRPBatch = " + batchNo + ") is null then 1 else ((select max(CRPRecNbr) from CRPPDAlloc where CRPBatch = " + batchNo + ") +1) end, " +
                         alloc[0].ToString() +", " + alloc[1].ToString() + ")";

                _jurisUtility.ExecuteNonQueryCommand(0, SQL);
            }
            
        }

        private void CreateBatchTrust(string batchNo)
        {
            Cursor.Current = Cursors.WaitCursor;
            toolStripStatusLabel.Text = "Creating Cash Receipt Batch Trust Record for Batch " + batchNo + "...";
            statusStrip.Refresh();
            UpdateStatus("Creating Cash Receipt  Batch Trust...", 5, 10);
            Application.DoEvents();

            string SQL = "select matsys, Bank_Code, Trust_Allocation from ##CombineChecks where Trust_Allocation is not null and Trust_Allocation <> 0 and batchNo = " + batchNo;

            DataSet trust = _jurisUtility.RecordsetFromSQL(SQL);

            foreach (DataRow alloc in trust.Tables[0].Rows)
            {
                SQL = "Insert into CRTrustAlloc(CRTBatch, CRTRecNbr, CRTSeqNbr, CRTMatter, CRTBank, CRTAmount) " +
                        "values (" + batchNo + ", case when(select max(CRTRecNbr) from CRTrustAlloc where CRTBatch = " + batchNo + ") is null then 1 else ((select max(CRTRecNbr) from CRTrustAlloc where CRTBatch = " + batchNo + ") +1) end, " +
                        "1," + alloc[0].ToString() + ", '" + alloc[1].ToString() + "', " + alloc[2].ToString() + ")";

                _jurisUtility.ExecuteNonQueryCommand(0, SQL);
            }

        }

        private void CreateBatchOther(string batchNo)
        {
            Cursor.Current = Cursors.WaitCursor;
            toolStripStatusLabel.Text = "Creating Cash Receipt Batch Non-Client Record for Batch " + batchNo + "...";
            statusStrip.Refresh();
            UpdateStatus("Creating Cash Receipt  Batch Non-Client...", 5, 10);
            Application.DoEvents();

            string SQL = "select Bank_Code, chartsysnbr, [Reference], Other_Allocation from ##CombineChecks where Other_Allocation is not null and Other_Allocation <> 0 and batchNo = " + batchNo;

            DataSet trust = _jurisUtility.RecordsetFromSQL(SQL);

            foreach (DataRow alloc in trust.Tables[0].Rows)
            {
                SQL = "Insert into CRNonCliAlloc(CRNBatch, CRNRecNbr, CRNSeqNbr, CRNBankCode, CRNCreditAccount, CRNReference, CRNAmount) " +
                        "values (" + batchNo + ", case when(select max(CRNRecNbr) from CRNonCliAlloc where CRNBatch = " + batchNo + ") is null then 1 else ((select max(CRNRecNbr) from CRNonCliAlloc where CRNBatch = " + batchNo + ") +1) end, " +
                        "1,'" + alloc[0].ToString() + "', " + alloc[1].ToString() + ", '" + alloc[2].ToString() + "', " + alloc[3].ToString() + ")";

                _jurisUtility.ExecuteNonQueryCommand(0, SQL);
            }


        }

        private void CreateBatchFees(string batchNo)
        {
            Cursor.Current = Cursors.WaitCursor;
            toolStripStatusLabel.Text = "Creating Cash Receipt Batch Fees Record for Batch " + batchNo + "...";
            statusStrip.Refresh();
            UpdateStatus("Creating Cash Receipt Batch Fees...", 6, 10);
            Application.DoEvents();

            string sql = "select Invoice_Number, matsys, sum(billFees) from ##CombineChecks where batchNo = " + batchNo + " and TType = 1 group by Invoice_Number, matsys having sum(billFees) <> 0";
            DataSet allocs = _jurisUtility.RecordsetFromSQL(sql);
            if (allocs != null && allocs.Tables.Count != 0 && allocs.Tables[0].Rows.Count != 0)
            {
                foreach (DataRow dd in allocs.Tables[0].Rows)
                {
                    double feeAllocsRemaining = Math.Round(Convert.ToDouble(dd[2].ToString()), 2);
                    sql = "SELECT [ARFTBillNbr]  ,[ARFTMatter] ,[ARFTTkpr],isnull([ARFTTaskCd], '') as ARFTTaskCd, isnull([ARFTActivityCd], '') as ARFTActivityCd, " +
                           " sum(arftactualamtbld - arftrcvd + arftadj) as prepost" + 
                            "  FROM [ARFTaskAlloc] " +
                            "   where arftbillnbr = " + dd[0].ToString() + " and arftmatter = " + dd[1].ToString() +
                             "  group by [ARFTBillNbr],[ARFTMatter],[ARFTTkpr],isnull([ARFTTaskCd], ''),isnull([ARFTActivityCd], '')";

                    DataSet arft = _jurisUtility.RecordsetFromSQL(sql);
                    List<FeeExpAlloc> feeExpList = new List<FeeExpAlloc>();
                    FeeExpAlloc fe;
                    foreach (DataRow ss in arft.Tables[0].Rows)
                    {
                        fe = new FeeExpAlloc();
                        fe.billNo = Convert.ToInt32(ss[0].ToString());
                        fe.mat = Convert.ToInt32(ss[1].ToString());
                        fe.tkpr = Convert.ToInt32(ss[2].ToString());
                        fe.amt = Math.Round(Convert.ToDouble(ss[5].ToString()), 2);
                        fe.code = ss[3].ToString();
                        fe.act = ss[4].ToString();
                        fe.pct = fe.amt / feeAllocsRemaining;
                        fe.allocAmt = Math.Round(feeAllocsRemaining * fe.pct, 2);
                        if (fe.allocAmt > fe.amt) // did we round up and to more than the total possible allocation? (force it if so)
                            fe.allocAmt = fe.amt;
                        feeExpList.Add(fe);
                    }

                    if (feeExpList.Select(x => x.allocAmt).Sum() != feeAllocsRemaining) //was there a rounding issue and the total allocations are  less than the actual total to be allocated??
                    { //its not possible to be more because we forced that earlier
                        double sum = feeExpList.Select(x => x.allocAmt).Sum();
                        double diff = feeAllocsRemaining - sum;
                        foreach (FeeExpAlloc ee in feeExpList)
                        {
                            if (ee.amt > ee.allocAmt) // assign the small leftover to whomever has room for it
                            {
                                double newAlloc = Math.Round(ee.amt - ee.allocAmt, 2);
                                if (newAlloc > diff) //if the difference is greater than what we have left to allocate, allocate whats left and exit the loop
                                {
                                    ee.allocAmt = ee.allocAmt + diff;
                                    diff = 0.00;
                                }
                                else
                                {
                                    ee.allocAmt = ee.allocAmt + newAlloc;
                                    diff = diff - newAlloc;
                                }
                            }
                            if (diff == 0) // we allocated everything, if not, next alloc
                                break;
                        }

                    }



                    foreach (FeeExpAlloc xx in feeExpList)
                    {
                        sql = "insert into #ARAlloc (matter, billnbr, tkpr, taskcd, actcd, expcd, exptype, allocAmt) values (" +
                            "" + xx.mat.ToString() + ", " + xx.billNo + ", " + xx.tkpr.ToString() + ", '" + xx.code + "', '" + xx.act + "', '', '', " + xx.allocAmt.ToString() + ")";
                        _jurisUtility.ExecuteNonQuery(0, sql);
                    }

                    //see if crfeealloc item exists first

                    string SQL = "Insert into CRFeeAlloc(crfbatch, crfrecnbr, crfmatter, crfbillnbr, crftkpr, crftaskcd, crfactivitycd, crfprepost, crfamount) " +
                        " Select  crbbatchnbr, crbreccount, matter, billnbr,tkpr, case when taskcd='' then null else taskcd end, case when actcd='' then null else actcd end, prepost, amt " +
                        " from (select matter, billnbr, tkpr, isnull(taskcd,'') as taskcd, isnull(actcd,'') as actcd, sum(allocAmt) as amt from #ARAlloc group by matter, billnbr, tkpr, isnull(taskcd,''), isnull(actcd,'')) AR " +
                        " Inner join (select arftmatter, arftbillnbr, arfttkpr, isnull(arfttaskcd,'') as ARFTTask, isnull(arftactivitycd,'') as ActivityCd, sum(arftactualamtbld - arftrcvd + arftadj) as Prepost " +
                        " from arftaskalloc group by arftmatter, arftbillnbr, arfttkpr, isnull(arfttaskcd,''), isnull(arftactivitycd,'')) ARM on matter=arftmatter and billnbr=arftbillnbr and arfttkpr=tkpr and taskcd=arfttask and actcd=activitycd, cashreceiptsbatch " +
                        " where crbbatchnbr=" + batchNo;
                    _jurisUtility.ExecuteNonQueryCommand(0, SQL);


                    SQL = "Update arftaskalloc set arftpend=crfamount " +
                        " from crfeealloc where crfbatch=" + batchNo + " and arftmatter=crfmatter and arftbillnbr=crfbillnbr and crftkpr=arfttkpr and isnull(crftaskcd,'')=isnull(arfttaskcd,'') " +
                        "  and isnull(crfactivitycd,'')=isnull(arftactivitycd,'')";
                    _jurisUtility.ExecuteNonQueryCommand(0, SQL);
                }
            }
            //else do nothing because there were no fee allocs

        }

        private void CreateBatchExps(string batchNo)
        {
            Cursor.Current = Cursors.WaitCursor;
            toolStripStatusLabel.Text = "Creating Cash Receipt Batch Exps Record for Batch " + batchNo + "...";
            statusStrip.Refresh();
            UpdateStatus("Creating Cash Receipt Batch Exps...", 8, 10);
            Application.DoEvents();

            //cash exp codes...
            string sql = "delete from #ARAlloc";
            _jurisUtility.ExecuteNonQuery(0, sql);
            //billCashExps, billNCashExps
            sql = "select Invoice_Number, matsys, sum(billCashExps) from ##CombineChecks where batchNo = " + batchNo + " and TType = 1 group by Invoice_Number, matsys having sum(billCashExps) <> 0";
            DataSet allocs = _jurisUtility.RecordsetFromSQL(sql);
            if (allocs != null && allocs.Tables.Count != 0 && allocs.Tables[0].Rows.Count != 0)
            {
                foreach (DataRow dd in allocs.Tables[0].Rows)
                {
                    double feeAllocsRemaining = Math.Round(Convert.ToDouble(dd[2].ToString()), 2);
                    sql = "SELECT [AREBillNbr]  ,[AREMatter] ,isnull([AREExpCd], '') as ARFTTaskCd, AREExpType, " +
                           " sum(AREBldAmount - ARERcvd + AREAdj) as prepost" +
                            "  FROM [ARExpAlloc] " +
                            "   where AREExpType = 'C' and AREBillNbr = " + dd[0].ToString() + " and AREMatter = " + dd[1].ToString() +
                             "  group by [AREBillNbr],[AREMatter],isnull([AREExpCd], ''),AREExpType";

                    DataSet arft = _jurisUtility.RecordsetFromSQL(sql);
                    List<FeeExpAlloc> feeExpList = new List<FeeExpAlloc>();
                    FeeExpAlloc fe;
                    foreach (DataRow ss in arft.Tables[0].Rows)
                    {
                        fe = new FeeExpAlloc();
                        fe.billNo = Convert.ToInt32(ss[0].ToString());
                        fe.mat = Convert.ToInt32(ss[1].ToString());
                        fe.tkpr = 0;
                        fe.amt = Math.Round(Convert.ToDouble(ss[4].ToString()), 2);
                        fe.code = ss[2].ToString();
                        fe.act = "";
                        fe.pct = fe.amt / feeAllocsRemaining;
                        fe.allocAmt = Math.Round(feeAllocsRemaining * fe.pct, 2);
                        if (fe.allocAmt > fe.amt) // did we round up and to more than the total possible allocation? (force it if so)
                            fe.allocAmt = fe.amt;
                        feeExpList.Add(fe);
                    }

                    if (feeExpList.Select(x => x.allocAmt).Sum() != feeAllocsRemaining) //was there a rounding issue and the total allocations are  less than the actual total to be allocated??
                    { //its not possible to be more because we forced that earlier
                        double sum = feeExpList.Select(x => x.allocAmt).Sum();
                        double diff = feeAllocsRemaining - sum;
                        foreach (FeeExpAlloc ee in feeExpList)
                        {
                            if (ee.amt > ee.allocAmt) // assign the small leftover to whomever has room for it
                            {
                                double newAlloc = Math.Round(ee.amt - ee.allocAmt, 2);
                                if (newAlloc > diff) //if the difference is greater than what we have left to allocate, allocate whats left and exit the loop
                                {
                                    ee.allocAmt = ee.allocAmt + diff;
                                    diff = 0.00;
                                }
                                else
                                {
                                    ee.allocAmt = ee.allocAmt + newAlloc;
                                    diff = diff - newAlloc;
                                }
                            }
                            if (diff == 0) // we allocated everything, if not, next alloc
                                break;
                        }

                    }
                    foreach (FeeExpAlloc xx in feeExpList)
                    {
                        sql = "insert into #ARAlloc (matter, billnbr, tkpr, taskcd, actcd, expcd, exptype, allocAmt) values (" +
                            "" + xx.mat.ToString() + ", " + xx.billNo + ", 0, '', '', '" + xx.code + "', 'C', " + xx.allocAmt.ToString() + ")";
                        _jurisUtility.ExecuteNonQuery(0, sql);
                    }
                    string SQL = "Insert into CRExpAlloc(crebatch, crerecnbr, crematter, crebillnbr, creexpcd, creexptype, creprepost, creamount) " +
                        " Select  crbbatchnbr, crbreccount, matter, billnbr,expcd, exptype, prepost, amt " +
                        " from (select matter, billnbr,expcd, exptype, sum(allocAmt) as amt from #ARAlloc  group by matter, billnbr,expcd, exptype) AR " +
                        " Inner join (select arematter, arebillnbr, areexpcd, areexptype, sum(arebldamount - arercvd + areadj) as Prepost " +
                        " from arexpalloc group by arematter, arebillnbr, areexpcd, areexptype) ARM on matter=arematter and billnbr=arebillnbr and areexpcd=expcd and areexptype=exptype, cashreceiptsbatch " +
                        " where crbbatchnbr=" + batchNo;
                    _jurisUtility.ExecuteNonQueryCommand(0, SQL);


                    SQL = "Update arexpalloc set arepend=creamount " +
                        " from crexpalloc where crebatch=" + batchNo + " and crematter=arematter and arebillnbr=crebillnbr and areexpcd=creexpcd and areexptype=creexptype ";
                    _jurisUtility.ExecuteNonQueryCommand(0, SQL);

                }



            }


            //same with exp codes...
            sql = "delete from #ARAlloc";
            _jurisUtility.ExecuteNonQuery(0, sql);
            //billCashExps, billNCashExps
            sql = "select Invoice_Number, matsys, sum(billNCashExps) from ##CombineChecks where batchNo = " + batchNo + " and TType = 1 group by Invoice_Number, matsys having sum(billNCashExps) <> 0";
            allocs.Clear();
            allocs = _jurisUtility.RecordsetFromSQL(sql);
            if (allocs != null && allocs.Tables.Count != 0 && allocs.Tables[0].Rows.Count != 0)
            {
                foreach (DataRow dd in allocs.Tables[0].Rows)
                {
                    double feeAllocsRemaining = Math.Round(Convert.ToDouble(dd[2].ToString()), 2);
                    sql = "SELECT [AREBillNbr]  ,[AREMatter] ,isnull([AREExpCd], '') as ARFTTaskCd, AREExpType, " +
                           " sum(AREBldAmount - ARERcvd + AREAdj) as prepost" +
                            "  FROM [ARExpAlloc] " +
                            "   where AREExpType = 'N' and AREBillNbr = " + dd[0].ToString() + " and AREMatter = " + dd[1].ToString() +
                             "  group by [AREBillNbr],[AREMatter],isnull([AREExpCd], ''),AREExpType";

                    DataSet arft = _jurisUtility.RecordsetFromSQL(sql);
                    List<FeeExpAlloc> feeExpList = new List<FeeExpAlloc>();
                    FeeExpAlloc fe;
                    foreach (DataRow ss in arft.Tables[0].Rows)
                    {
                        fe = new FeeExpAlloc();
                        fe.billNo = Convert.ToInt32(ss[0].ToString());
                        fe.mat = Convert.ToInt32(ss[1].ToString());
                        fe.tkpr = 0;
                        fe.amt = Math.Round(Convert.ToDouble(ss[4].ToString()), 2);
                        fe.code = ss[2].ToString();
                        fe.act = "";
                        fe.pct = fe.amt / feeAllocsRemaining;
                        fe.allocAmt = Math.Round(feeAllocsRemaining * fe.pct, 2);
                        if (fe.allocAmt > fe.amt) // did we round up and to more than the total possible allocation? (force it if so)
                            fe.allocAmt = fe.amt;
                        feeExpList.Add(fe);
                    }

                    if (feeExpList.Select(x => x.allocAmt).Sum() != feeAllocsRemaining) //was there a rounding issue and the total allocations are  less than the actual total to be allocated??
                    { //its not possible to be more because we forced that earlier
                        double sum = feeExpList.Select(x => x.allocAmt).Sum();
                        double diff = feeAllocsRemaining - sum;
                        foreach (FeeExpAlloc ee in feeExpList)
                        {
                            if (ee.amt > ee.allocAmt) // assign the small leftover to whomever has room for it
                            {
                                double newAlloc = Math.Round(ee.amt - ee.allocAmt, 2);
                                if (newAlloc > diff) //if the difference is greater than what we have left to allocate, allocate whats left and exit the loop
                                {
                                    ee.allocAmt = ee.allocAmt + diff;
                                    diff = 0.00;
                                }
                                else
                                {
                                    ee.allocAmt = ee.allocAmt + newAlloc;
                                    diff = diff - newAlloc;
                                }
                            }
                            if (diff == 0) // we allocated everything, if not, next alloc
                                break;
                        }

                    }
                    foreach (FeeExpAlloc xx in feeExpList)
                    {
                        sql = "insert into #ARAlloc (matter, billnbr, tkpr, taskcd, actcd, expcd, exptype, allocAmt) values (" +
                            "" + xx.mat.ToString() + ", " + xx.billNo + ", 0, '', '', '" + xx.code + "', 'N', " + xx.allocAmt.ToString() + ")";
                        _jurisUtility.ExecuteNonQuery(0, sql);
                    }
                    string SQL = "Insert into CRExpAlloc(crebatch, crerecnbr, crematter, crebillnbr, creexpcd, creexptype, creprepost, creamount) " +
                        " Select  crbbatchnbr, crbreccount, matter, billnbr,expcd, exptype, prepost, amt " +
                        " from (select matter, billnbr,expcd, exptype, sum(allocAmt) as amt from #ARAlloc  group by matter, billnbr,expcd, exptype) AR " +
                        " Inner join (select arematter, arebillnbr, areexpcd, areexptype, sum(arebldamount - arercvd + areadj) as Prepost " +
                        " from arexpalloc group by arematter, arebillnbr, areexpcd, areexptype) ARM on matter=arematter and billnbr=arebillnbr and areexpcd=expcd and areexptype=exptype, cashreceiptsbatch " +
                        " where crbbatchnbr=" + batchNo;
                    _jurisUtility.ExecuteNonQueryCommand(0, SQL);


                    SQL = "Update arexpalloc set arepend=creamount " +
                        " from crexpalloc where crebatch=" + batchNo + " and crematter=arematter and arebillnbr=crebillnbr and areexpcd=creexpcd and areexptype=creexptype ";
                    _jurisUtility.ExecuteNonQueryCommand(0, SQL);

                }



            }
            sql = "delete from #ARAlloc";
            _jurisUtility.ExecuteNonQuery(0, sql);
            sql = "delete from ##CombineChecks";
            _jurisUtility.ExecuteNonQuery(0, sql);

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


        public string PYear = "";
        public string PNbr = "";
        public string DOrder = "";
        private System.Drawing.Point pt;

        string EmpInt = "";
        string EmpSys = "";
        bool codeIsNumericClient = false;
        int lengthOfCodeClient = 0;
        bool codeIsNumericMatter = false;
        int lengthOfCodeMatter = 0;


        private void button1_Click(object sender, EventArgs e)
        {
            DoDaFix();
        }

        private void buttonReport_Click(object sender, EventArgs e)
        {
            checks.Clear();
            errors.Clear();
            selectedSpreadsheet = false;
            OpenFileDialog dlg = new OpenFileDialog();
            double totalAllocations = 0.00;
            if (dlg.ShowDialog() == DialogResult.OK)
            {
                string fileName;
                fileName = dlg.FileName.ToString();
                System.Data.OleDb.OleDbConnection MyConnection;
                System.Data.DataSet DtSet;
                System.Data.OleDb.OleDbDataAdapter MyCommand;
                MyConnection = new System.Data.OleDb.OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + fileName + ";Extended Properties='Excel 8.0;HDR=YES;';");

                MyCommand = new System.Data.OleDb.OleDbDataAdapter("select * from [Sheet1$] where [CheckAmount]<>0", MyConnection);

                DtSet = new System.Data.DataSet();
                MyCommand.Fill(DtSet);

                MyConnection.Close();
                Check ck;
                DataError error;



                if (DtSet != null && DtSet.Tables.Count != 0 && DtSet.Tables[0].Rows.Count != 0)
                {

                        int row = 1;
                        foreach (DataRow dr in DtSet.Tables[0].Rows) // go through each row of the spreadsheet
                        {
                            row++;
                            int clisys = 0;
                            int matsys = 0;
                            ck = new Check();
                            if (isReceiptType(dr[0].ToString().Trim().ToUpper())) //proper receipt type?
                                ck.receiptType = dr[0].ToString().Trim().ToUpper();
                            else
                            {
                                error = new DataError();
                                error.rowNum = row;
                                error.error = "The Receipt Type is not valid. Valid receipt types are 'A', 'P', 'T', 'X'";
                                errors.Add(error);
                                continue;
                            }
                            if (!IsDate(dr[1].ToString().Trim())) //good date?
                            {
                                error = new DataError();
                                error.rowNum = row;
                                error.error = "The Deposit Date is not a valid date";
                                errors.Add(error);
                                continue;
                            }
                            else
                                ck.depDate = dr[1].ToString().Trim();
                            if (string.IsNullOrEmpty(dr[2].ToString().Trim()) || dr[2].ToString().Trim().Length > 10)
                            {
                                error = new DataError();
                                error.rowNum = row;
                                error.error = "The Check Number is required and must be 10 characters or less";
                                errors.Add(error);
                                continue;
                            }
                            ck.checkNum = dr[2].ToString().Trim();
                            if (!IsNumeric(dr[3].ToString().Replace("$", "").Replace(",", "").Trim())) // is it a money value?
                            {
                                error = new DataError();
                                error.rowNum = row;
                                error.error = "The Check Amount is not a valid money value ($ and commas are allowed)";
                                errors.Add(error);
                                continue;
                            }
                            else
                                ck.checkAmt = Convert.ToDouble(dr[3].ToString().Replace("$", "").Replace(",", "").Trim());
                            if (!IsDate(dr[4].ToString().Trim()))
                            {
                                error = new DataError();
                                error.rowNum = row;
                                error.error = "The Check Date is not a valid date";
                                errors.Add(error);
                                continue;
                            }
                            else
                                ck.checkDate = dr[4].ToString().Trim();
                            if (string.IsNullOrEmpty(dr[5].ToString().Trim()) || dr[5].ToString().Trim().Length > 50)
                            {
                                error = new DataError();
                                error.rowNum = row;
                                error.error = "Payor is required on all checks and must be 50 characters or less";
                                errors.Add(error);
                                continue;
                            }
                            else
                                ck.payor = dr[5].ToString().Trim();
                            char switchType = Convert.ToChar(ck.receiptType);
                            switch (switchType)
                            {
                                case 'A': //AR alloc
                                    clisys = verifyClientCode(dr[6].ToString().Trim()); //real client?
                                    if (clisys != 0)
                                    {
                                        ck.client = dr[6].ToString().Trim();
                                        ck.clisysnbr = clisys;
                                    }
                                    else
                                    {
                                        error = new DataError();
                                        error.rowNum = row;
                                        error.error = "The Client Code is not valid. It must match exactly as it displays in Juris";
                                        errors.Add(error);
                                        continue;
                                    }
                                    matsys = verifyMatterCode(dr[7].ToString().Trim(), clisys); // real matter?
                                    if (matsys != 0)
                                    {
                                        ck.matter = dr[7].ToString().Trim();
                                        ck.matsysnbr = matsys;
                                    }
                                    else
                                    {
                                        error = new DataError();
                                        error.rowNum = row;
                                        error.error = "The Matter Code is not valid. It must match exactly as it displays in Juris";
                                        errors.Add(error);
                                        continue;
                                    }
                                    if (string.IsNullOrEmpty(dr[12].ToString().Trim()) || !IsNumeric(dr[12].ToString().Trim())) //required...not valid record without it
                                    {
                                        error = new DataError();
                                        error.rowNum = row;
                                        error.error = "The Invoice Number is required and must be nuemric (integer)";
                                        errors.Add(error);
                                        continue;
                                    }
                                    else
                                        ck.invNumber = Convert.ToInt32(dr[12].ToString().Trim());
                                    if (string.IsNullOrEmpty(dr[11].ToString().Trim()) || string.IsNullOrEmpty(dr[13].ToString().Trim()) || string.IsNullOrEmpty(dr[14].ToString().Trim())) // there is no inv date, bill total or bill balance...pull it from data via billnbr
                                    {

                                        Check inv = new Check();
                                        inv = getInvDetails(ck.invNumber, ck.matsysnbr);
                                        if (inv.invNumber == 0) // we didnt get any data on that inv number
                                        {
                                            error = new DataError();
                                            error.rowNum = row;
                                            error.error = "The Invoice Number: " + ck.invNumber.ToString() + " does not appear to be valid or have any allocations accociated with it";
                                            errors.Add(error);
                                            continue;
                                        }
                                        else
                                        {
                                            ck.invDate = inv.invDate;
                                            ck.billTotal = inv.billTotal;
                                            ck.billBalance = inv.billBalance;
                                        }

                                    }
                                    else // they populated it
                                    {
                                        if (IsDate(dr[11].ToString().Trim()))
                                            ck.invDate = dr[11].ToString().Trim();
                                        else
                                        {
                                            error = new DataError();
                                            error.rowNum = row;
                                            error.error = "The Invoice Date is not a valid date";
                                            errors.Add(error);
                                            continue;
                                        }
                                        if (!IsNumeric(dr[13].ToString().Replace("$", "").Replace(",", "").Trim())) // is it a money value?
                                        {
                                            error = new DataError();
                                            error.rowNum = row;
                                            error.error = "The Bill Total Amount is not a valid money value ($ and commas are allowed)";
                                            errors.Add(error);
                                            continue;
                                        }
                                        else
                                            ck.billTotal = Convert.ToDouble(dr[13].ToString().Replace("$", "").Replace(",", "").Trim());
                                        if (!IsNumeric(dr[14].ToString().Replace("$", "").Replace(",", "").Trim())) // is it a money value?
                                        {
                                            error = new DataError();
                                            error.rowNum = row;
                                            error.error = "The Bill Balance Amount is not a valid money value ($ and commas are allowed)";
                                            errors.Add(error);
                                            continue;
                                        }
                                        else
                                            ck.billBalance = Convert.ToDouble(dr[14].ToString().Replace("$", "").Replace(",", "").Trim());



                                    }
                                    if (!IsNumeric(dr[15].ToString().Replace("$", "").Replace(",", "").Trim())) // is it a money value?
                                    {
                                        error = new DataError();
                                        error.rowNum = row;
                                        error.error = "The AR Amount is not a valid money value ($ and commas are allowed)";
                                        errors.Add(error);
                                        continue;
                                    }
                                    else
                                    {
                                        ck.allocationAmount = Convert.ToDouble(dr[15].ToString().Replace("$", "").Replace(",", "").Trim());
                                        totalAllocations = totalAllocations + ck.allocationAmount;
                                        ck.ar = Convert.ToDouble(dr[15].ToString().Replace("$", "").Replace(",", "").Trim());
                                    }
                                    break;
                                case 'P': // PPD alloc
                                    clisys = verifyClientCode(dr[6].ToString().Trim()); //real client?
                                    if (clisys != 0)
                                    {
                                        ck.client = dr[6].ToString().Trim();
                                        ck.clisysnbr = clisys;
                                    }
                                    else
                                    {
                                        error = new DataError();
                                        error.rowNum = row;
                                        error.error = "The Client Code is not valid. It must match exactly as it displays in Juris";
                                        errors.Add(error);
                                        continue;
                                    }
                                    matsys = verifyMatterCode(dr[7].ToString().Trim(), clisys); // real matter?
                                    if (matsys != 0)
                                    {
                                        ck.matter = dr[7].ToString().Trim();
                                        ck.matsysnbr = matsys;
                                    }
                                    else
                                    {
                                        error = new DataError();
                                        error.rowNum = row;
                                        error.error = "The Matter Code is not valid. It must match exactly as it displays in Juris";
                                        errors.Add(error);
                                        continue;
                                    }
                                    if (!IsNumeric(dr[16].ToString().Replace("$", "").Replace(",", "").Trim())) // is it a money value?
                                    {
                                        error = new DataError();
                                        error.rowNum = row;
                                        error.error = "The PPD Amount is not a valid money value ($ and commas are allowed)";
                                        errors.Add(error);
                                        continue;
                                    }
                                    else
                                    {
                                        ck.allocationAmount = Convert.ToDouble(dr[16].ToString().Replace("$", "").Replace(",", "").Trim());
                                        totalAllocations = totalAllocations + ck.allocationAmount;
                                        ck.ppd = Convert.ToDouble(dr[16].ToString().Replace("$", "").Replace(",", "").Trim());
                                    }
                                    break;
                                case 'T': //Trust alloc
                                    clisys = verifyClientCode(dr[6].ToString().Trim()); //real client?
                                    if (clisys != 0)
                                    {
                                        ck.client = dr[6].ToString().Trim();
                                        ck.clisysnbr = clisys;
                                    }
                                    else
                                    {
                                        error = new DataError();
                                        error.rowNum = row;
                                        error.error = "The Client Code is not valid. It must match exactly as it displays in Juris";
                                        errors.Add(error);
                                        continue;
                                    }
                                    matsys = verifyMatterCode(dr[7].ToString().Trim(), clisys); // real matter?
                                    if (matsys != 0)
                                    {
                                        ck.matter = dr[7].ToString().Trim();
                                        ck.matsysnbr = matsys;
                                    }
                                    else
                                    {
                                        error = new DataError();
                                        error.rowNum = row;
                                        error.error = "The Matter Code is not valid. It must match exactly as it displays in Juris";
                                        errors.Add(error);
                                        continue;
                                    }
                                    if (!IsNumeric(dr[17].ToString().Replace("$", "").Replace(",", "").Trim())) // is it a money value?
                                    {
                                        error = new DataError();
                                        error.rowNum = row;
                                        error.error = "The Trust Amount is not a valid money value ($ and commas are allowed)";
                                        errors.Add(error);
                                        continue;
                                    }
                                    else
                                    {
                                        ck.allocationAmount = Convert.ToDouble(dr[17].ToString().Replace("$", "").Replace(",", "").Trim());
                                        totalAllocations = totalAllocations + ck.allocationAmount;
                                        ck.trust = Convert.ToDouble(dr[17].ToString().Replace("$", "").Replace(",", "").Trim());
                                }
                                    if (isValidBank(dr[8].ToString().Trim()))
                                        ck.bankCode = dr[8].ToString().Trim();
                                    else
                                    {
                                        error = new DataError();
                                        error.rowNum = row;
                                        error.error = "The Bank Code is not valid. It must match exactly as it displays in Juris";
                                        errors.Add(error);
                                        continue;
                                    }
                                    if (!isTrustAcctSetUpForMatter(ck.bankCode, ck.matsysnbr))
                                    {
                                    error = new DataError();
                                    error.rowNum = row;
                                    error.error = "The matter is not set up with a trust account. It must be set up before hand";
                                    errors.Add(error);
                                    continue;

                                }
                                    break;
                                case 'X': //Other (non cli) alloc
                                    if (!IsNumeric(dr[18].ToString().Replace("$", "").Replace(",", "").Trim())) // is it a money value?
                                    {
                                        error = new DataError();
                                        error.rowNum = row;
                                        error.error = "The Other Amount is not a valid money value ($ and commas are allowed)";
                                        errors.Add(error);
                                        continue;
                                    }
                                    else
                                    {
                                        ck.allocationAmount = Convert.ToDouble(dr[18].ToString().Replace("$", "").Replace(",", "").Trim());
                                        totalAllocations = totalAllocations + ck.allocationAmount;
                                        ck.noncli = Convert.ToDouble(dr[18].ToString().Replace("$", "").Replace(",", "").Trim());
                                }
                                    if (isValidBank(dr[8].ToString().Trim()))
                                        ck.bankCode = dr[8].ToString().Trim();
                                    else
                                    {
                                        error = new DataError();
                                        error.rowNum = row;
                                        error.error = "The Bank Code is not valid. It must match exactly as it displays in Juris";
                                        errors.Add(error);
                                        continue;
                                    }
                                    int chartsysnbr = isValidGL(dr[9].ToString().Trim());
                                    if (chartsysnbr != 0)
                                    {
                                        ck.glAccount = dr[9].ToString().Trim();
                                        ck.chartsysnbr = chartsysnbr;
                                    }
                                    else
                                    {
                                        error = new DataError();
                                        error.rowNum = row;
                                        error.error = "The GL Account is not valid. It must match exactly as it displays in Juris";
                                        errors.Add(error);
                                        continue;
                                    }
                                    if (!string.IsNullOrEmpty(dr[10].ToString().Trim()) || dr[10].ToString().Trim().Length > 250)
                                        ck.reference = dr[10].ToString().Trim();
                                    else
                                    {
                                        error = new DataError();
                                        error.rowNum = row;
                                        error.error = "The Reference field is required for all 'Other items' (Receipt Type 'X') and must be 250 characters or less";
                                        errors.Add(error);
                                        continue;
                                    }
                                    break;



                            }
                            ck.rowNum = row;
                            checks.Add(ck);

                        }

                        //check to see if payors match across all check numbers and if allocations match check total and if check amount is same for each check number
                        string sql = "IF OBJECT_ID('tempdb.dbo.##CheckTest', 'U') IS NOT NULL DROP TABLE [dbo].[##CheckTest]";
                        _jurisUtility.ExecuteNonQuery(0, sql);
                        sql = "IF OBJECT_ID('tempdb.dbo.##valid', 'U') IS NOT NULL DROP TABLE [dbo].[##valid]";
                        _jurisUtility.ExecuteNonQuery(0, sql);
                        sql = "create table ##CheckTest (checkNum varchar(30), checkAmt money, [payor] varchar(1500), Alloc money, row int, [validation] int)";
                        _jurisUtility.ExecuteNonQuery(0, sql);
                        foreach (Check check in checks)
                        {
                            sql = "insert into ##CheckTest (checkNum, checkAmt, [payor], Alloc, row, [validation]) values ('" + check.checkNum + "', " + check.checkAmt.ToString() + ", '" + check.payor + "', " + check.allocationAmount.ToString() + ", " + check.rowNum.ToString() + ", 1)";
                            _jurisUtility.ExecuteNonQuery(0, sql);
                        }
                        sql = "select llj.checkNum, llj.valid into ##valid from  (select checkNum, sum([validation]) as valid from ##CheckTest group by checkNum) llj";
                        _jurisUtility.ExecuteNonQuery(0, sql);

                        

                        validateCheckTotalsPerCheck();
                        validatePayorsPerCheckNumber();
                        validateAllocationsMatchCheckAmount();

                        //if any record for a check fails...all records for the check fail
                        invalidateAllRecordsFoFailedCheck();

                    //remove allocation amounts for bad checks from list and temp tables as they arent used any more
                    foreach (Check cc in checks.Where(x => x.isError == true))
                    {
                        totalAllocations = totalAllocations - cc.allocationAmount;
                        sql = "delete from ##CheckTest where checkNum = '" + cc.checkNum + "'";
                        _jurisUtility.ExecuteNonQuery(0, sql);
                        sql = "delete from ##valid where checkNum = '" + cc.checkNum + "'";
                        _jurisUtility.ExecuteNonQuery(0, sql);
                    }
                    //do we have any good checks left?
                    int numOfGoodChecks = checks.Where(x => x.isError == false).Count();

                //get check amounts and make sure they match the allocation totals for remaining good checks
                    if (totalAllocations == getTotalAllocsFromCheck() && numOfGoodChecks > 0)
                    {
                        //combine into one row per check
                        sql = "IF OBJECT_ID('tempdb.dbo.##CombineChecks', 'U') IS NOT NULL DROP TABLE [dbo].[##CombineChecks]";
                        _jurisUtility.ExecuteNonQuery(0, sql);
                        sql = "create table ##CombineChecks (Deposit_Date dateTime, Check_Number varchar(30), Check_Amount decimal(15,2), Check_Date datetime, [Payor] varchar(1500), " +
                               " Client_Code varchar(20), Matter_Code varchar(20), Bank_Code varchar(8), GL_Account varchar(20), [Reference] varchar(1500), " +
                               " Invoice_Date datetime, Invoice_Number int, Bill_Total decimal(15,2), Bill_Balance decimal(15,2), AR_Allocation decimal(15,2), Fees decimal(15,2), Cash_Exps decimal(15,2), " +
                               " NonCash_Exps decimal(15,2), Taxes1 decimal(15,2), Taxes2 decimal(15,2), Taxes3 decimal(15,2), Surcharge decimal(15,2), " +
                               " Interest decimal(15,2), PPD_Allocation decimal(15,2), Trust_Allocation decimal(15,2), Other_Allocation decimal(15,2), TType int, batchNo int, matsys int, " +
                               "billFees decimal(15,2), billCashExps decimal(15,2), billNCashExps decimal(15,2), billInt decimal(15,2), billTax1 decimal(15,2), billTax2 decimal(15,2), billTax3 decimal(15,2), billSur decimal(15,2), chartsysnbr int)";
                        _jurisUtility.ExecuteNonQuery(0, sql);
                        foreach (Check check in checks.Where(x => x.isError == false))
                        {
                            sql = "insert into ##CombineChecks (Deposit_Date, Check_Number, Check_Amount, Check_Date, [Payor], " +
                               " Client_Code, Matter_Code, Bank_Code, GL_Account, [Reference], " +
                               " Invoice_Date, Invoice_Number, Bill_Total, Bill_Balance, AR_Allocation, Fees, Cash_Exps, NonCash_Exps, Taxes1, " +
                               " Taxes2, Taxes3, Surcharge, Interest, PPD_Allocation, Trust_Allocation, Other_Allocation, TTYpe, batchNo, matsys, " +
                               " billFees, billCashExps, billNCashExps, billInt, billTax1, billTax2, billTax3, billSur, chartsysnbr) " +
                                " values (convert(varchar,'" + check.depDate + "', 101), '" + check.checkNum + "', " + check.checkAmt.ToString() + ", convert(varchar, '" + check.checkDate + "', 101), '" + check.payor + "', " +
                                " '" + check.client +"', '" + check.matter + "', '" + check.bankCode + "', '" + check.glAccount + "', '" + check.reference +"', " +
                                " convert(varchar, '" + check.invDate + "', 101), " + check.invNumber.ToString() + ", " + check.billTotal.ToString() + ", " + check.billBalance.ToString() + ", " +
                                check.ar.ToString() + ", 0.00, 0.00, 0.00,0.00,0.00,0.00,0.00,0.00," + check.ppd.ToString() + ", " + check.trust.ToString() + ", " + check.noncli.ToString() + ", "   +
                                " case when '" + check.receiptType + "' = 'A' then 1 else 2 end, 0, " + check.matsysnbr + ", 0.00,0.00,0.00,0.00,0.00,0.00,0.00,0.00, " + check.chartsysnbr.ToString() + ")";
                            _jurisUtility.ExecuteNonQuery(0, sql);
                        }
                        //  get the total amount due from each category so they cant over allocate (like if they have 400 in fees and 100 in cahs exps left, we dont want them allocating
                        //500 in fees as its more than the fees owed
                        sql = "update ##CombineChecks set billFees =([ARMFeeBld] - [ARMFeeRcvd] + [ARMFeeAdj]), billCashExps = ([ARMCshExpBld] - [ARMCshExpRcvd] + [ARMCshExpAdj]), " +
                            " billNCashExps = ([ARMNCshExpBld] - [ARMNCshExpRcvd] + [ARMNCshExpAdj]), billInt = ([ARMIntBld] - [ARMIntRcvd] + [ARMIntAdj]), " +
                            " billTax1 = ([ARMTax1Bld] - [ARMTax1Rcvd] + [ARMTax1Adj] ), billTax2 = ([ARMTax2Bld] - [ARMTax2Rcvd] + [ARMTax2Adj] ), billTax3 = ([ARMTax3Bld] - [ARMTax3Rcvd]+ [ARMTax3Adj]), " +
                            " billSur = ([ARMSurchgBld] - [ARMSurchgRcvd] + [ARMSurchgAdj]) " +
                            " from ##CombineChecks inner join armatalloc on matsys = armmatter and Invoice_Number = armbillnbr";
                        _jurisUtility.ExecuteNonQuery(0, sql);

                        //get data and display
                        sql = "select Check_Number, Client_Code, Matter_Code, " +
                               " Invoice_Date, Invoice_Number, " +
                                " max(Check_Date) as Check_Date, max([Payor]) as [Payor], max(Check_Amount) as Check_Amount, max(Deposit_Date) as Deposit_Date, " +
                                " max(Bill_Balance) as Bill_Balance, max(Bill_Total) as Bill_Total, " +
                                " sum(AR_Allocation) as AR_Allocation, sum(Fees) as Fees, sum(Cash_Exps) as Cash_Exps, " +
                               "  sum(NonCash_Exps) as NonCash_Exps, sum(Taxes1) as Taxes1, sum(Taxes2) as Taxes2, " +
                                " sum(Taxes3) as Taxes3, sum(Surcharge) as Surcharge, sum(Interest) as Interest, " +
                                " sum(billFees) as billFees, sum(billCashExps) as billCashExps, sum(billNCashExps) as billNCashExps, sum(billInt) as billInt, " +
                                " sum(billTax1) as billTax1, sum(billTax2) as billTax2, sum(billTax3) as billTax3, sum(billSur) as billSur, matsys " +
                                " from ##CombineChecks " +
                                " where TType = 1 " +
                                " group by Check_Number, Client_Code, Matter_Code,  " +
                                " Invoice_Date, Invoice_Number, matsys";

                        DataSet AR = _jurisUtility.RecordsetFromSQL(sql);

                        sql = "select Check_Number, Client_Code, Matter_Code, Bank_Code, GL_Account, [Reference],  " +
                                " max(Check_Date) as Check_Date, max([Payor]) as [Payor], max(Check_Amount) as Check_Amount, max(Deposit_Date) as Deposit_Date, " +
                                " sum(PPD_Allocation) as PPD_Allocation, sum(Trust_Allocation) as Trust_Allocation, sum(Other_Allocation) as Other_Allocation " +
                                " from ##CombineChecks " +
                                " where TType = 2 " +
                                " group by Check_Number, Client_Code, Matter_Code, Bank_Code, GL_Account, [Reference],   " +
                                " Invoice_Date, Invoice_Number";

                        DataSet others = _jurisUtility.RecordsetFromSQL(sql);

                        ReportDisplay rt = new ReportDisplay(AR.Tables[0], others.Tables[0], _jurisUtility); 
                        rt.ShowDialog();
                        selectedSpreadsheet = true;
                        button1.Enabled = true;
                        buttonReport.Enabled = false;
                        button1.ForeColor = System.Drawing.Color.Red;
                    }
                    else
                    {
                        error = new DataError();
                        error.rowNum = 0;
                        error.error = "The allocation total does not match the check total for all checks. Please verify the allocation totals match the check totals";
                        errors.Add(error);

                        DialogResult DR = MessageBox.Show("There was at least one error in the spreadsheet. All errors must be corrected before moving forward. Do you want to view the errors?", "Allocation mismatch confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                        if (DR == DialogResult.Yes)
                        {
                            sql = "IF OBJECT_ID('tempdb.dbo.##ErrorsTool', 'U') IS NOT NULL DROP TABLE [dbo].[##ErrorsTool]";
                            _jurisUtility.ExecuteNonQuery(0, sql);
                            sql = "create table ##ErrorsTool (row int, err varchar(500))";
                            _jurisUtility.ExecuteNonQuery(0, sql);
                            foreach (DataError err in errors)
                            {
                                sql = "insert into ##ErrorsTool (row, err) values(" + err.rowNum.ToString() + ", '" + err.error + "')";
                                _jurisUtility.ExecuteNonQuery(0, sql);
                            }
                            sql = "select row as Row_Number, err as Error_Message from ##ErrorsTool";
                            DataSet DSErr = new DataSet();
                            DSErr = _jurisUtility.RecordsetFromSQL(sql);
                            if (DSErr != null && DSErr.Tables.Count > 0 && DSErr.Tables[0].Rows.Count > 0)
                            {

                                DataTable tabErr = DSErr.Tables[0];
                                ErrorDisplay ed = new ErrorDisplay(tabErr);
                                ed.ShowDialog();

                            }

                            else
                                MessageBox.Show("No Errors");
                        }
                        selectedSpreadsheet = false;
                    }
                    
                }
                else
                {
                    MessageBox.Show("Your spreadsheet does not appear to contain any data" + "\r\n" + "Please update the spreadsheet", "Data Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                

            }
        }

        public static DataTable ToDataTable<T>(List<T> items)
        {
            DataTable dataTable = new DataTable(typeof(T).Name);

            //Get all the properties
            PropertyInfo[] Props = typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance);
            foreach (PropertyInfo prop in Props)
            {
                //Defining type of data column gives proper data table 
                var type = (prop.PropertyType.IsGenericType && prop.PropertyType.GetGenericTypeDefinition() == typeof(Nullable<>) ? Nullable.GetUnderlyingType(prop.PropertyType) : prop.PropertyType);
                //Setting column names as Property names
                dataTable.Columns.Add(prop.Name, type);
            }
            foreach (T item in items)
            {
                var values = new object[Props.Length];
                for (int i = 0; i < Props.Length; i++)
                {
                    //inserting property values to datatable rows
                    values[i] = Props[i].GetValue(item, null);
                }
                dataTable.Rows.Add(values);
            }
            //put a breakpoint here and check datatable
            return dataTable;
        }

        private void invalidateAllRecordsFoFailedCheck()
        {
            List<string> badChecks = new List<string>();
            foreach (Check check in checks)
            {
                if (check.isError)
                    badChecks.Add(check.checkNum);
            }

            foreach (string bc in badChecks)
            {
                foreach (Check cc in checks.Where(w => w.checkNum.Equals(bc)))
                    cc.isError = true;
            }
        }

        private void validateCheckTotalsPerCheck()
        {
            DataError error;
            string sql = "select ##CheckTest.checkNum, count(checkAmt) as CT from ##CheckTest inner join ##valid on ##CheckTest.checkNum = ##valid.checkNum group by ##CheckTest.checkNum, ##valid.valid having count(checkAmt) <> valid ";//do any checks have a different cheeck amount (like check 4567 having 2 records
            DataSet ff = _jurisUtility.RecordsetFromSQL(sql);                                                           //and 1 saying check amt is 50.00 and the other saying its 52.00
            if (ff != null && ff.Tables.Count > 0 && ff.Tables[0].Rows.Count > 0)
            {
                foreach (DataRow dr in ff.Tables[0].Rows)
                {
                    sql = "select top 1 row from ##CheckTest where checkNum = '" + dr[0].ToString() + "'"; //get row number for each check so we can display it to customer and log error
                    DataSet dd = _jurisUtility.RecordsetFromSQL(sql);
                    if (dd != null && dd.Tables.Count > 0 && dd.Tables[0].Rows.Count > 0)
                    {
                        foreach (DataRow dr1 in dd.Tables[0].Rows)
                        {
                            error = new DataError();
                            error.rowNum = Convert.ToInt32(dr1[0].ToString());
                            error.error = "There are conflicting check amounts for the records of check number " + dr[0].ToString() + ". All check amounts must be the same for all records of the same check number";
                            errors.Add(error);
                            foreach (Check cc in checks.Where(x => x.checkNum.Equals(dr[0].ToString())))
                            {
                                cc.isError = true;
                            }
                        }
                    }
                }

            }

        }

        private void validatePayorsPerCheckNumber()
        {
            DataError error;
            string sql = "select ##CheckTest.checkNum, count(payor) as CT from ##CheckTest inner join ##valid on ##CheckTest.checkNum = ##valid.checkNum group by ##CheckTest.checkNum, ##valid.valid having count(payor) <> valid  ";//all records with the same check number need to have the same payor
            DataSet ff = _jurisUtility.RecordsetFromSQL(sql);                                                          
            if (ff != null && ff.Tables.Count > 0 && ff.Tables[0].Rows.Count > 0)
            {
                foreach (DataRow dr in ff.Tables[0].Rows)
                {
                    sql = "select top 1 row from ##CheckTest where checkNum = '" + dr[0].ToString() + "'"; //get row number for each check so we can display it to customer and log error
                    DataSet dd = _jurisUtility.RecordsetFromSQL(sql);
                    if (dd != null && dd.Tables.Count > 0 && dd.Tables[0].Rows.Count > 0)
                    {
                        foreach (DataRow dr1 in dd.Tables[0].Rows)
                        {
                            error = new DataError();
                            error.rowNum = Convert.ToInt32(dr1[0].ToString());
                            error.error = "There are conflicting payors for check number " + dr[0].ToString() + ". All payors must be the sme for records with the same check number";
                            errors.Add(error);
                            foreach (Check cc in checks.Where(x => x.checkNum.Equals(dr[0].ToString())))
                            {
                                cc.isError = true;
                            }
                        }
                    }
                }

            }

        }

        private void validateAllocationsMatchCheckAmount()
        {
            DataError error;
            string sql = "select checkNum, checkAmt, sum(Alloc) as CT from ##CheckTest group by checkNum, checkAmt having sum(Alloc) <> checkAmt ";//all allocations shold add up to check amount for each check number
            DataSet ff = _jurisUtility.RecordsetFromSQL(sql);
            if (ff != null && ff.Tables.Count > 0 && ff.Tables[0].Rows.Count > 0)
            {
                foreach (DataRow dr in ff.Tables[0].Rows)
                {
                    sql = "select top 1 row from ##CheckTest where checkNum = '" + dr[0].ToString() + "'"; //get row number for each check so we can display it to customer and log error
                    DataSet dd = _jurisUtility.RecordsetFromSQL(sql);
                    if (dd != null && dd.Tables.Count > 0 && dd.Tables[0].Rows.Count > 0)
                    {
                        foreach (DataRow dr1 in dd.Tables[0].Rows)
                        {
                            error = new DataError();
                            error.rowNum = Convert.ToInt32(dr1[0].ToString());
                            error.error = "The total allocations do not match the check amount for the records of check number " + dr[0].ToString() + ". The sum of allocations must equal the check amount for any given check number";
                            errors.Add(error);
                            foreach (Check cc in checks.Where(x => x.checkNum.Equals(dr[0].ToString())))
                            {
                                cc.isError = true;
                            }
                        }
                    }
                }

            }

        }

        private double getTotalAllocsFromCheck()
        {
            double totalCheckAmount = 0.00;
            string sql = "select checkNum, sum(checkAmt) as CA from (select distinct checkNum, checkAmt from ##CheckTest) lla group by checkNum ";//total of check amt should equal the total allocs (we send check amt back for verification)
            DataSet ff = _jurisUtility.RecordsetFromSQL(sql);
            if (ff != null && ff.Tables.Count > 0 && ff.Tables[0].Rows.Count > 0)
            {
                foreach (DataRow dr in ff.Tables[0].Rows)
                {
                    totalCheckAmount = Convert.ToDouble(dr[1].ToString());
                }

            }

            return totalCheckAmount;

        }

        private void cbUser_SelectedIndexChanged(object sender, EventArgs e)
        {
            EmpInt = this.cbUser.GetItemText(this.cbUser.SelectedItem).Split(' ')[0];

            string SqlEmp = "select empsysnbr from employee where empinitials='" + EmpInt + "'";
            DataSet dt = _jurisUtility.RecordsetFromSQL(SqlEmp);

            DataTable db = dt.Tables[0];

            if (db.Rows.Count == 0)
            { EmpSys = "1"; }
            else
            {
                EmpSys = db.Rows[0]["empsysnbr"].ToString();
            }
        }

        private bool isTrustAcctSetUpForMatter(string bank, int matsys)
        {
            bool isGood = false;
            string sql = "select * from trustaccount where TAMatter = " + matsys.ToString() + " and TABank = '" + bank + "'";
            DataSet tbank = _jurisUtility.RecordsetFromSQL(sql);
            if (tbank != null && tbank.Tables.Count != 0 && tbank.Tables[0].Rows.Count != 0)
                isGood = true;
            return isGood;
        }

        private void labelDescription_Click(object sender, EventArgs e)
        {

        }

        private bool isReceiptType(string ttype)
        {
            if (string.IsNullOrEmpty(ttype))
                return false;
            else if (ttype.Equals("A") || ttype.Equals("P") || ttype.Equals("T") || ttype.Equals("X"))
                return true;
            else
                return false;
        }

        private Check getInvDetails(int invNum, int matsys)
        {
            Check cck = new Check();
            cck.invNumber = 0;
            string sql = " select sum(armbaldue) as ARBal, sum([ARMFeeBld] + ARMCshExpBld + ARMNCshExpBld + ARMSurchgBld + ARMTax1Bld + ARMTax3Bld + ARMTax2Bld + ARMIntBld) as TotalBilled, " +
                             " arbilldate, arbillnbr " +
                             " from ARMatAlloc " +
                             " inner join arbill on arbillnbr = armbillnbr " +
                             " where arbillnbr = " + invNum.ToString() + " and armmatter = " + matsys.ToString() +
                             " group by arbilldate, arbillnbr ";
            DataSet LH = _jurisUtility.RecordsetFromSQL(sql);
            if (LH != null && LH.Tables.Count != 0 && LH.Tables[0].Rows.Count != 0)
            {
                foreach (DataRow dr in LH.Tables[0].Rows) // go through each row of the spreadsheet
                {
                    cck.invNumber = invNum;
                    cck.invDate = dr[2].ToString().Trim();
                    cck.billTotal = Convert.ToDouble(dr[1].ToString().Trim());
                    cck.billBalance = Convert.ToDouble(dr[0].ToString().Trim());
                }
            }

            return cck;
        }

        private int verifyClientCode(string clicode)
        {
            int clisys = 0;
            string code = formatClientCode(clicode);
            string sql = "select clisysnbr from client where clicode = '" + code + "'";
            DataSet dds = _jurisUtility.RecordsetFromSQL(sql);
            if (dds != null && dds.Tables.Count > 0)
            {
                foreach (DataRow dr in dds.Tables[0].Rows) //client already exists
                {
                    clisys = Convert.ToInt32(dr[0].ToString());
                }

            }


            return clisys;
        }

        private int verifyMatterCode(string clicode, int clisys)
        {

            int matsys = 0;
            string code = formatMatterCode(clicode);
            string sql = "select matsysnbr from matter inner join client on clisysnbr = matclinbr where matcode = '" + code + "' and clisysnbr = " + clisys.ToString();
            DataSet dds = _jurisUtility.RecordsetFromSQL(sql);
            if (dds != null && dds.Tables.Count > 0)
            {
                foreach (DataRow dr in dds.Tables[0].Rows) //matter already exists
                {
                    matsys = Convert.ToInt32(dr[0].ToString());
                }

            }


            return matsys;
        }

        private string formatClientCode(string code)
        {
            string formattedCode = "";
            if (codeIsNumericClient)
            {
                formattedCode = "000000000000" + code;
                formattedCode = formattedCode.Substring(formattedCode.Length - 12, 12);
            }
            else
                formattedCode = code;
            return formattedCode;

        }

        private string formatMatterCode(string code)
        {
            string formattedCode = "";
            if (codeIsNumericMatter)
            {
                formattedCode = "000000000000" + code;
                formattedCode = formattedCode.Substring(formattedCode.Length - 12, 12);
            }
            else
                formattedCode = code;
            return formattedCode;

        }

        private void getSettings()
        {
            //matter
            string sysparam = "  select SpTxtValue from sysparam where SpName = 'FldMatter'";
            DataSet dds2 = _jurisUtility.RecordsetFromSQL(sysparam);
            string cell = "";
            if (dds2 != null && dds2.Tables.Count > 0)
            {
                foreach (DataRow dr in dds2.Tables[0].Rows)
                {
                    cell = dr[0].ToString();
                }

            }
            string[] test = cell.Split(',');
            lengthOfCodeMatter = Convert.ToInt32(test[2]);

            if (test[1].Equals("C"))
                codeIsNumericMatter = false;
            else
                codeIsNumericMatter = true;

            //client
            sysparam = "  select SpTxtValue from sysparam where SpName = 'FldClient'";
            dds2.Clear();
            dds2 = _jurisUtility.RecordsetFromSQL(sysparam);
            if (dds2 != null && dds2.Tables.Count > 0)
            {
                foreach (DataRow dr in dds2.Tables[0].Rows)
                {
                    cell = dr[0].ToString();
                }

            }
            string[] test1 = cell.Split(',');
            lengthOfCodeClient = Convert.ToInt32(test1[2]);

            if (test1[1].Equals("C"))
                codeIsNumericClient = false;
            else
                codeIsNumericClient = true;

        }

        private bool isValidBank(string bCode)
        {

            string sql = "select * from bankaccount where BnkCode = '" + bCode + "'";
            DataSet dds = _jurisUtility.RecordsetFromSQL(sql);
            if (dds != null && dds.Tables.Count > 0 && dds.Tables[0].Rows.Count != 0)
            {
                foreach (DataRow dr in dds.Tables[0].Rows) 
                {
                    return true;
                }
                return false;
            }
            else
                return false;

        }

        private int isValidGL(string glCode)
        {
            int chtsys = 0;
            string sql = "select chtsysnbr from chartofaccounts where dbo.jfn_FormatChartOfAccount(ChartOfAccounts.ChtSysNbr) = '" + glCode + "'";
            DataSet dds = _jurisUtility.RecordsetFromSQL(sql);
            if (dds != null && dds.Tables.Count > 0 && dds.Tables[0].Rows.Count != 0)
            {
                foreach (DataRow dr in dds.Tables[0].Rows)
                {
                    chtsys = Convert.ToInt32(dr[0].ToString());
                }

            }
            else
                chtsys = 0;
            return chtsys;
        }

        private void radioButtonExcel_CheckedChanged(object sender, EventArgs e)
        {
            buttonReport.Enabled = radioButtonExcel.Checked;
            this.Hide();
            ManualEntry me = new ManualEntry(_jurisUtility, this.Location, Convert.ToInt32(EmpSys), EmpInt);
            me.ShowDialog();
            //check to see if they added anything then bring main form back
            button1.Enabled = true;
        }

        private void radioButtonManual_CheckedChanged(object sender, EventArgs e)
        {
            pt = this.Location;
        }
    }
}
