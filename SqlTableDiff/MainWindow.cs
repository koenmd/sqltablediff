using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Diagnostics;

// TODO: Add buttons to allow the user to test each connection (Source and Destination).
// TODO: Add UI elements for unused properties of Diff.Processor.
// TODO: Finish the feature that merges all outputted sql files into one file (default to OFF).
// TODO: Instead of distributing tablediff.exe, make a config setting that points to it's location.
// Alternatively, make a batch file that copies it from it's location so that the end user can
// bring it with them wherever they want.

namespace SqlTableDiff
{
    public partial class MainWindow : Form
    {

        private System.Data.SqlClient.SqlConnection _connectionSrc = null;

        public MainWindow()
        {
            InitializeComponent();

            LoadDefaults();

            OutFolderDlg.SelectedPath = OutFolderTxt.Text;

            SrcAuthBinding.DataSource = SrcAuthSqlOpt;
            DestAuthBinding.DataSource = DestAuthSqlOpt;
        }

        #region Default Values

        private void LoadDefaults()
        {
            LoadDefaultsSource();
            LoadDefaultsDestination();
            LoadDefaultsOutput();
            LoadDefaultsTables();
            LoadDefaultsOptions();
        }

        void LoadDefaultsSource()
        {
            var val = Defaults.Source.Default;

            SrcServerTxt.Text = val.Server;
            if (val.AuthType.Trim().ToLower() == Diff.DbAuthType.Windows.ToString().ToLower())
                SrcAuthWinOpt.Checked = true;
            else
                SrcAuthSqlOpt.Checked = true;
            SrcUserTxt.Text = val.AuthUser;
            SrcPassTxt.Text = val.AuthPass;
            SrcDatabaseTxt.Text = val.DbName;
            SrcSchemaTxt.Text = val.DbSchema;
        }

        void LoadDefaultsDestination()
        {
            var val = Defaults.Destination.Default;

            DestServerTxt.Text = val.Server;

            if (val.AuthType.Trim().ToLower() == Diff.DbAuthType.Windows.ToString().ToLower())
                DestAuthWinOpt.Checked = true;
            else
                DestAuthSqlOpt.Checked = true;
            DestUserTxt.Text = val.AuthUser;
            DestPassTxt.Text = val.AuthPass;
            DestDatabaseTxt.Text = val.DbName;
            DestSchemaTxt.Text = val.DbSchema;
        }

        void LoadDefaultsOutput()
        {
            var val = Defaults.Output.Default;

            var folder = System.IO.Path.Combine(App.ExeFolderLocation, val.Folder);
            System.Diagnostics.Debug.Print("Output Folder: " + folder);
            OutFolderTxt.Text = folder;

            OutScriptSingleFileChk.Checked = val.SingleScript;
        }

        void LoadDefaultsTables()
        {
            var val = Defaults.Tables.Default;

            // Turn the comma delimited list of tables into a NewLine delimited list.
            var tables = val.DefaultList
                .Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(r => r.Trim())
                .Aggregate((agg, r) => agg + Environment.NewLine + r);

            TablesTxt.Text = tables;
        }

        void LoadDefaultsOptions()
        {
            var val = Defaults.AdvancedOptions.Default;

            OptColCompareChk.Checked = val.ColumCompare;
            OptStrictSchemaChk.Checked = val.SchemaStrict;
            OptFastCompareChk.Checked = val.FastCompare;

            OptLargeObjBytesTxt.Text = val.LargeObjectBytes;
            OptConnectTimeoutTxt.Text = val.ConnectTimeout;
            OptRetryCountTxt.Text = val.RetryCount;
            OptRetryIntervalTxt.Text = val.RetryInterval;
        }

        #endregion

        #region Menu

        void FileExitItm_Click(object sender, EventArgs e)
        {
            Close();
        }

        void HelpAboutItm_Click(object sender, EventArgs e)
        {
            MessageBox.Show(
@"Sql Table Diff Utility 1.0

Author(s): Wayne Bloss", "About", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        void HelpInstructItm_Click(object sender, EventArgs e)
        {
            MessageBox.Show(
@"          Sql Table Diff Utility 1.0

This tool will show the differences in a list of tables between
a source and destination database and generate the SQL necessary
to update the destination database.

1) Enter the Source and Destination Sql Server information.

2) Edit the list of tables to be compared if necessary. The default
list may be changed as noted below.

3) Consider using the Advanced Options if a previous run resulted
in an error or other problem.

4) Press the GO! button. You can observe the progress of each table
in the text box below the GO! button.

5) After execution or at any time during execution, you may press the
Open Output Folder button to see the generated SQL scripts and the
report file. To open the report file in the default text editor, hit
the Open Report File button.

(NOTE: The default values for all information on the main window may be
changed by editing the SqlTableDiff.exe.config file.)

",
            "Help", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        #endregion

        #region Output

        private void OutFolderBtn_Click(object sender, EventArgs e)
        {
            if (OutFolderDlg.ShowDialog(this) != System.Windows.Forms.DialogResult.OK) return;
            OutFolderTxt.Text = OutFolderDlg.SelectedPath;
        }

        private void OutFolderOpenBtn_Click(object sender, EventArgs e)
        {
            if (_currentOutputFolder != null)
                System.Diagnostics.Process.Start(_currentOutputFolder);
            else
            {
                System.IO.Directory.CreateDirectory(OutFolderTxt.Text);
                System.Diagnostics.Process.Start(OutFolderTxt.Text);
            }
        }

        private void OutReportOpenBtn_Click(object sender, EventArgs e)
        {
            if (_currentReportFile == null) return;
            System.Diagnostics.Process.Start(_currentReportFile);
        }

        #endregion

        #region Execution

        string _currentOutputFolder;
        string _currentReportFile;

        bool _isExecuting;

        void ExecBtn_Click(object sender, EventArgs e)
        {
            if (_isExecuting)
                CancelExecution();
            else
                Execute();
        }

        private void ExecPrint()
        {
            ExecPrint("");
        }

        void ExecPrint(string value)
        {
            ExecInfo.AppendText(value + Environment.NewLine);
        }

        private void ExecPrintSectionEnd()
        {
            ExecPrint("---------------------------------------------------------------------");
            ExecPrint();
        }

        void Execute()
        {
            try
            {
                OnBeforeExecute();
                PrepareOutputFolder();
                Process();
            }
            catch (Exception ex)
            {
                ExecPrint("Error: " + ex.Message);
            }
            finally
            {
                OnAfterExecute();
            }
        }

        void OnBeforeExecute()
        {
            _isExecuting = true;
            Cursor = Cursors.WaitCursor;
            ExecInfo.Clear();
            StatusLbl.Text = "Running...";
            ExecBtn.Text = "&Cancel";
        }

        void OnAfterExecute()
        {
            var cancelled = _cancelExecution;
            _isExecuting = false;
            _cancelExecution = false;

            StatusLbl.Text = "Completed";
            ExecBtn.Text = "&GO!";

            if (cancelled)
                ExecPrint("Cancelled by user.");
            else
                ExecPrint("Done.");

            Cursor = Cursors.Default;
        }

        #endregion

        #region Processing

        void PrepareOutputFolder()
        {
            var now = DateTime.Now;
            var dateName = now.Year.ToString("0000") + "_" + now.Month.ToString("00") + "_" + now.Day.ToString("00") +
                "-" + now.Hour.ToString("00") + "_" + now.Minute.ToString("00") + "_" + now.Second.ToString("00");

            var folder = System.IO.Path.Combine(OutFolderTxt.Text, dateName);

            _currentOutputFolder = folder;
            ExecPrint("Outputing to " + folder);
            ExecPrintSectionEnd();
            System.IO.Directory.CreateDirectory(_currentOutputFolder);
        }

        void Process()
        {
            var proc = new Diff.Processor();
            var src = proc.Source;
            src.Server = SrcServerTxt.Text.Trim();
            src.AuthType = GetDbAuthType(SrcAuthWinOpt);
            src.AuthUser = SrcUserTxt.Text.Trim();
            src.AuthPass = SrcPassTxt.Text.Trim();
            src.Name = SrcDatabaseTxt.Text.Trim();
            src.Schema = SrcSchemaTxt.Text.Trim();

            var dest = proc.Destination;
            dest.Server = DestServerTxt.Text.Trim();
            dest.AuthType = GetDbAuthType(DestAuthWinOpt);
            dest.AuthUser = DestUserTxt.Text.Trim();
            dest.AuthPass = DestPassTxt.Text.Trim();
            dest.Name = DestDatabaseTxt.Text.Trim();
            dest.Schema = DestSchemaTxt.Text.Trim();

            proc.OutputFolder = _currentOutputFolder;

            var opt = proc.Options;
            opt.ColumnCompare = OptColCompareChk.Checked;
            opt.StrictSchema = OptStrictSchemaChk.Checked;
            opt.FastCompare = OptFastCompareChk.Checked;
            opt.LargeObjectBytes = GetNullableInt(OptLargeObjBytesTxt);
            opt.ConnectTimeout = GetNullableInt(OptConnectTimeoutTxt);
            opt.RetryCount = GetNullableInt(OptRetryCountTxt);
            opt.RetryInterval = GetNullableInt(OptRetryIntervalTxt);

            ProcessTables(proc);
        }

        void ProcessTables(Diff.Processor proc)
        {
            var list = TablesTxt.Text.Split(new string[] { Environment.NewLine }, StringSplitOptions.RemoveEmptyEntries);

            foreach (var item in list)
            {
                if (Cancelled()) return;
                ExecPrint("Table " + item + " being processed...");
                ExecPrint();

                proc.TableName = item;
                proc.Execute();
                _currentReportFile = proc.ReportFile;

                ExecPrint(proc.StandardOutput);
                ExecPrint("Table " + item + " result: " + proc.ExitStatus);
                ExecPrintSectionEnd();
            }
        }

        Diff.DbAuthType GetDbAuthType(RadioButton winAuthOption)
        {
            return winAuthOption.Checked ? Diff.DbAuthType.Windows : Diff.DbAuthType.Sql;
        }

        int? GetNullableInt(TextBox box)
        {
            var txt = box.Text.Trim();
            if (txt.Length > 0)
            {
                int result;
                if (int.TryParse(txt, out result)) return result;
            }
            return null;
        }

        #endregion

        #region Cancellation

        bool _cancelExecution;

        void CancelExecution()
        {
            if (_cancelExecution) return; // (already cancelled.)
            ExecPrint("Cancelling...");
            _cancelExecution = true;
        }

        bool Cancelled()
        {
            Application.DoEvents();
            return _cancelExecution;
        }

        #endregion

        private void btnConnect_Click(object sender, EventArgs e)
        {
            if (_connectionSrc != null)
            {
                try
                {
                    if (_connectionSrc.State == ConnectionState.Open)
                    {
                        _connectionSrc.Close();
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Exception", buttons: MessageBoxButtons.OK, icon: MessageBoxIcon.Error);
                }
                finally
                {
                    _connectionSrc.Dispose();
                    _connectionSrc = null;
                }
            }

            Trace.WriteLine(GetConnectionString());
            _connectionSrc = new System.Data.SqlClient.SqlConnection(GetConnectionString());
            _connectionSrc.Open();

            RefreshDatabaseList();
            RefreshTableList();

        }

        private void RefreshTableList()
        {
            if (_connectionSrc != null && _connectionSrc.State == ConnectionState.Open)
            {
                DataTable dt = _connectionSrc.GetSchema("Tables", new string[] { null, null, null, "BASE TABLE" });
                DataView view = new DataView(dt);
                view.Sort = "TABLE_TYPE ASC, TABLE_SCHEMA ASC, TABLE_NAME ASC";

                if (view.Count > 0)
                {
                    clbTables.Items.Clear();
                    foreach (DataRowView row in view)
                    {
                        if (row["TABLE_TYPE"].ToString() == "BASE TABLE")
                        {
                            if (row["TABLE_SCHEMA"].ToString() == SrcSchemaTxt.Text)
                            {
                                clbTables.Items.Add(row["TABLE_NAME"].ToString());
                            }
                        }
                    }

                }

            }
        }

        private void RefreshDatabaseList()
        {
            if (_connectionSrc != null && _connectionSrc.State == ConnectionState.Open)
            {
                DataTable dt = _connectionSrc.GetSchema("Databases");
                DataView view = new DataView(dt);
                view.Sort = "database_name ASC";
                if(view.Count > 0)
                {
                    SrcCboDatabases.Items.Clear();
                    foreach(DataRowView row in view)
                    {
                        SrcCboDatabases.Items.Add(row["database_name"].ToString());
                    }
                }
            }
        }

        private string GetConnectionString()
        {
            List<Exception> lstExceptions = new List<Exception>();
            if (!string.IsNullOrWhiteSpace(SrcServerTxt.Text))
            {
                if (SrcAuthWinOpt.Checked)
                {
                    return string.Format("Integrated Security=SSPI;Persist Security Info=True;Initial Catalog={0};Data Source={1}", SrcDatabaseTxt.Text, SrcServerTxt.Text);
                }
                if (SrcAuthSqlOpt.Checked)
                {
                    if (!string.IsNullOrWhiteSpace(SrcUserTxt.Text) && !string.IsNullOrWhiteSpace(SrcPassTxt.Text))
                    {
                        return string.Format("Persist Security Info=True;User ID={0};Password={1};Initial Catalog={2};Data Source={3};", SrcUserTxt.Text, SrcPassTxt.Text, SrcDatabaseTxt.Text, SrcServerTxt.Text);
                    }
                }

            }
            else
            {
                Exception ex = new Exception("Fill out server name please");
                lstExceptions.Add(ex);
            }

            if (lstExceptions != null && lstExceptions.Count > 0)
            {
                foreach (Exception ex in lstExceptions)
                    throw ex;
            }

            return "";
        }

        private void SrcServerTxt_TextChanged(object sender, EventArgs e)
        {

        }

        private void MainWindow_Load(object sender, EventArgs e)
        {

        }

        private void MainWindow_FormClosing(object sender, FormClosingEventArgs e)
        {
            if(_connectionSrc != null)
            {
                if (_connectionSrc.State == ConnectionState.Open)
                {
                    _connectionSrc.Close();
                }
                _connectionSrc.Dispose();
                _connectionSrc = null;
            }
        }
    }
}
