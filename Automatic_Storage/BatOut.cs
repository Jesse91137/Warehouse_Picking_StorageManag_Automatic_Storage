using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Security;
using System.Text;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Threading;
using System.Runtime.InteropServices;

namespace Automatic_Storage
{
    public partial class BatOut : Form
    {
        private OpenFileDialog FileUpload1;
        //private Button selectButton;
        DataSet dsData = new DataSet();
        public DataTable table;
        DataTable cpDataTable = new DataTable();
        string OutDate = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
        StringBuilder sb = new StringBuilder();
        bool txtFlag = true;
        //1→ 批次1 ; 2→批次2
        private static string btn_sender = string.Empty;
        #region 視窗ReSize
        int X = new int();  //窗口寬度
        int Y = new int(); //窗口高度
        float fgX = new float(); //寬度縮放比例
        float fgY = new float(); //高度縮放比例
        bool isLoaded;  // 是否已設定各控制的尺寸資料到Tag屬性
        #endregion

        #region 子視窗回傳值
        private bool childValueButton;
        public bool MsgFromChildButton
        {
            set { childValueButton = value; }
        }
        private string childValuePosition;
        public string MsgFromChildPosition
        {
            set { childValuePosition = value; }
        }
        private string childValueEngSr;
        public string MsgFromChildEngSr
        {
            set { childValueEngSr = value; }
        }
        //↓↓原本chk="OK" 變數不異動直接使用變成Amount領用數量
        public string MsgFromChildChk
        {
            set { childValueChk = value; }
        }
        
        private string childValueChk;
        
        private string childValueItemE;
        public string MsgFromChildItemE
        {
            set { childValueItemE = value; }
        }
        private string childValueAmount;
        public string MsgFromChildAmount
        {
            set { childValueAmount = value; }
        }

        #endregion


        public class BatItem : IComparable<BatItem>, IEquatable<BatItem>
        {
            public int sno { get; set; }
            public string Item1 { get; set; }
            public string Station { get; set; }
            public string Spec { get; set; }
            public string Item2 { get; set; }
            public string NA { get; set; }
            public string Chk { get; set; }
            public string Position { get; set; }
            public string Engsr { get; set; }
            public string Memo { get; set; }

            public override string ToString()
            {
                return "sno: " + sno + "   Item1: " + Item1 + "Station: " + Station + "   Spec: " + Spec + "Item2: " + Item2 + "   NA: " + NA +
                    "Chk: " + Chk + "   Position: " + Position + "Engsr:" + Engsr + "Memo: " + Memo;
            }
            public override bool Equals(object obj)
            {
                if (obj == null) return false;
                BatItem objAsPart = obj as BatItem;
                if (objAsPart == null) return false;
                else return Equals(objAsPart);
            }
            public int SortByNameAscending(string name1, string name2)
            {
                return name1.CompareTo(name2);
            }

            // Default comparer for Part type.
            public int CompareTo(BatItem comparePart)
            {
                // A null value means that this object is greater.
                if (comparePart == null)
                    return 1;

                else
                    return this.Position.CompareTo(comparePart.Position);
            }
            public override int GetHashCode()
            {
                return sno;
            }
            public bool Equals(BatItem other)
            {
                if (other == null) return false;
                return (this.sno.Equals(other.sno));
            }
            // Should also override == and != operators.
        }
        private DataTable columnsData()
        {
            using (DataTable table = new DataTable())
            {
                switch (btn_sender)
                {
                    case "1":
                        // Add columns.
                        table.Columns.Add("sno", typeof(string));
                        table.Columns.Add("process", typeof(string));
                        table.Columns.Add("itemE", typeof(string));
                        table.Columns.Add("itemC", typeof(string));
                        table.Columns.Add("spec", typeof(string));
                        //table.Columns.Add("position", typeof(string));
                        table.Columns.Add("storageP", typeof(string));
                        table.Columns.Add("count", typeof(string));
                        table.Columns.Add("storagePC", typeof(string));
                        table.Columns.Add("KM10", typeof(string));
                        //table.Columns.Add("received", typeof(string));
                        //table.Columns.Add("postCnt", typeof(string));
                        //table.Columns.Add("surplus", typeof(string));
                        table.Columns.Add("replace", typeof(string));
                        //table.Columns.Add("mark", typeof(string));
                        //table.Columns.Add("F15", typeof(string));
                        //table.Columns.Add("F16", typeof(string));
                        //table.Columns.Add("diff", typeof(string));
                        table.Columns.Add("chk", typeof(string));
                        break;
                    case "2":
                        // Add columns.
                        table.Columns.Add("sno", typeof(string));
                        table.Columns.Add("process", typeof(string));
                        table.Columns.Add("itemE", typeof(string));
                        table.Columns.Add("itemC", typeof(string));
                        table.Columns.Add("spec", typeof(string));                        
                        table.Columns.Add("count", typeof(string));                                                
                        table.Columns.Add("chk", typeof(string));
                        break;
                    default:
                        break;
                }                
                return table;
            }
        }
        public BatOut()
        {
            InitializeComponent();
            isLoaded = false;
            
        }        
        private void BatOut_Load(object sender, EventArgs e)
        {            
            #region 獲取窗體info
            X = this.Width;//獲取窗體的寬度
                Y = this.Height;//獲取窗體的高度
                isLoaded = true;// 已設定各控制項的尺寸到Tag屬性中
                SetTag(this);//調用方法
            #endregion

            FileUpload1 = new OpenFileDialog()
            {
                FileName = "Select a text file",
                Filter = "Text files (*.xls)|*.xls",
                Title = "Open text file"
            };
        }

        private void selectButton_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(txt_wono.Text))
            {
                txtFlag = false;
                MessageBox.Show("請先輸入工單號碼");
                return;
            }
            if(txtFlag)
            {
                if (FileUpload1.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        var fileName = FileUpload1.FileName;

                        FileInfo _info = new FileInfo(fileName);
                        string _new = Application.StartupPath + "\\Upload\\" + _info.Name;
                        if (File.Exists(fileName))
                        {
                            _info.CopyTo(_new, true);
                            txt_path.Text = _new;
                        }
                        //dataBind();
                    }
                    catch (SecurityException ex)
                    {
                        MessageBox.Show($"Security error.\n\nError message: {ex.Message}\n\n" +
                        $"Details:\n\n{ex.StackTrace}");
                    }
                }
            }
        }        
        List<BatItem> ls = new List<BatItem>();
        bool msgResponse = false;
        
        private void commitButton_Click(object sender, EventArgs e)
        {
            //this.loadExcelBW.WorkerSupportsCancellation = true; //允許中斷
            //this.loadExcelBW.RunWorkerAsync(); //呼叫背景程式            
            btn_sender = "1";
            table = table ?? columnsData();
            LoadExcel();
        }
        private void commitButton2_Click(object sender, EventArgs e)
        {
            btn_sender = "2";
            table = table ?? columnsData();
            LoadExcel();
        }
        private void LoadExcel()
        {
            //Form.CheckForIllegalCrossThreadCalls = false;
            if (FileUpload1.FileName != "")
            {
                string path = Application.StartupPath + "~/Upload/";
                //string fileName = Path.Combine(path, Path.GetFileName(FileUpload1.PostedFile.FileName));
                string fileName = txt_path.Text;
                string fileExtension = System.IO.Path.GetExtension(FileUpload1.FileName).ToLower();
                string[] allowExtension = { ".xls" };
                bool fileOk = false;
                //判斷文件類型
                for (int i = 0; i < allowExtension.Length; i++)
                {
                    if (fileExtension == allowExtension[i])
                    {
                        fileOk = true;
                        break;
                    }
                }
                if (fileOk)
                {

                    if (dsRE.Tables[0].Rows.Count > 0)
                    {
                        if (MessageBox.Show("此工單有之前出庫紀錄是否讀取紀錄", "", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                        {
                            msgResponse = true;
                        }
                    }
                    //FileUpload1.PostedFile.SaveAs(fileName);
                    try
                    {
                        string excelString = "";
                        #region office版本
                        if (fileExtension == ".xls" || fileExtension == ".XLS")
                        {
                            #region office 97-2003
                            excelString = "Provider=Microsoft.Jet.OLEDB.4.0;" +
                                                                 "Data Source=" + fileName + ";" +
                                                                 "Extended Properties='Excel 8.0;HDR=No;IMEX=1\'";
                            #endregion
                        }
                        if (fileExtension == ".xlsx" || fileExtension == ".XLSX" || fileExtension == ".xlsm" || fileExtension == ".XLSM")
                        {
                            #region office 2007
                            excelString = "Provider=Microsoft.ACE.OLEDB.12.0;" +
                                                                 "Data Source=" + fileName +
                                                                 ";Extended Properties='Excel 12.0 Xml;HDR=YES;IMEX=1';";
                            #endregion
                        }
                        #endregion
                        OleDbConnection cnn = new OleDbConnection(excelString);
                        cnn.Open();
                        OleDbCommand cmd = new OleDbCommand();
                        cmd.Connection = cnn;
                        OleDbDataAdapter adapter = new OleDbDataAdapter();
                        switch (btn_sender)
                        {
                            case "1":
                                cmd.CommandText = " SELECT F1,F2,F3,F4,F6,F7,F8,F9,F13,'' as chk FROM [Sheet1$] where F1 is not null or F1 <> '' ";
                                break;
                            case "2":
                                cmd.CommandText = " SELECT F2,F3,F4,F6,F7, '' as chk FROM [Sheet1$] ";
                                break;
                            default:
                                break;
                        }
                        
                        adapter.SelectCommand = cmd;
                        dsData = new DataSet();
                        adapter.Fill(dsData);
                        cnn = null;
                        cmd = null;
                        adapter = null;
                        string sqlinput = string.Empty;
                        string sqldetail = string.Empty;
                        string sqlchkFirst = string.Empty;
                        int sno = 0;
                        //itemE = 昶料號, itemC = 客料號, spec = 規格, chk = 確認, storageP = 發料倉,
                        //count = 應領數/需求量, storagePC = 發料倉庫存, process = 製程, replace= 替代群組
                        //KM10 = KM10, received = 已領數量, surplus = 結餘 , mark = 備註 , diff = 用料差異 , position = 儲位 , postCnt = 發料數造
                        string itemE = "", itemC = "", spec = "", chk = "", storageP = "", position = ""
                            , count = "", storagePC = "", process = "", replace = "", KM10 = "", countED = ""
                            , surplus = "", mark = "", diff = "";
                        string finddata = "";
                        string strShift = "";
                        progressBar1.Minimum = 0;
                        progressBar1.Maximum = dsData.Tables[0].Rows.Count-1;
                        progressBar1.Step = 1;

                        switch (btn_sender)
                        {
                            case "1":
                                //從第二筆開始                        
                                for (int i = 2; i < dsData.Tables[0].Rows.Count; i++)
                                {
                                    DataRow dr = table.NewRow();
                                    dr["sno"] = i;
                                    dr["process"] = dsData.Tables[0].Rows[i]["F1"].ToString();
                                    dr["itemE"] = dsData.Tables[0].Rows[i]["F2"].ToString();
                                    dr["itemC"] = dsData.Tables[0].Rows[i]["F3"].ToString();
                                    dr["spec"] = dsData.Tables[0].Rows[i]["F4"].ToString();
                                    //dr["position"] = dsData.Tables[0].Rows[i]["F5"].ToString();
                                    dr["storageP"] = dsData.Tables[0].Rows[i]["F6"].ToString();
                                    dr["count"] = dsData.Tables[0].Rows[i]["F7"].ToString();
                                    dr["storagePC"] = dsData.Tables[0].Rows[i]["F8"].ToString();
                                    dr["KM10"] = dsData.Tables[0].Rows[i]["F9"].ToString();
                                    dr["replace"] = dsData.Tables[0].Rows[i]["F13"].ToString();
                                    dr["chk"] = dsData.Tables[0].Rows[i]["chk"].ToString();
                                    table.Rows.Add(dr);
                                }
                                break;
                            case "2":
                                //從第二筆開始                        
                                for (int i = 3; i < dsData.Tables[0].Rows.Count; i++)
                                {
                                    DataRow dr = table.NewRow();
                                    dr["sno"] = i;
                                    dr["process"] = dsData.Tables[0].Rows[i]["F2"].ToString();
                                    dr["itemE"] = dsData.Tables[0].Rows[i]["F3"].ToString();
                                    dr["itemC"] = dsData.Tables[0].Rows[i]["F4"].ToString();
                                    dr["spec"] = dsData.Tables[0].Rows[i]["F6"].ToString();                                    
                                    dr["count"] = dsData.Tables[0].Rows[i]["F7"].ToString();                                    
                                    dr["chk"] = dsData.Tables[0].Rows[i]["chk"].ToString();
                                    table.Rows.Add(dr);
                                }
                                break;
                            default:
                                break;
                        }
                        

                        //回寫
                        if (msgResponse)
                        {
                            foreach (DataRow row in dsRE.Tables[0].Rows)
                            {
                                var rowsToUpdate =
                                    table.AsEnumerable()
                                    .Where(r => r.Field<string>("itemE") == row["Item_No_Master"].ToString().Trim().ToUpper());
                                foreach (var rows in rowsToUpdate)
                                {
                                    rows.SetField("chk", row["amount"]);
                                }
                            }
                        }

                        dataGridView1.DataSource = null;
                        dataGridView1.DataSource = table;
                    }
                    catch (Exception ee)
                    {
                        throw;
                    }
                }
                else
                {
                    //lblMessage.Text = "文件格式類型不對,只支援Excel文件!";
                }
            }
        }
        public DataTable ConvertToDataTable<T>(IList<T> data)
        {            
            PropertyDescriptorCollection properties =
               TypeDescriptor.GetProperties(typeof(T));
            DataTable table = new DataTable();
            foreach (PropertyDescriptor prop in properties)
                table.Columns.Add(prop.Name, Nullable.GetUnderlyingType(prop.PropertyType) ?? prop.PropertyType);
            foreach (T item in data)
            {
                DataRow row = table.NewRow();
                foreach (PropertyDescriptor prop in properties)
                    row[prop.Name] = prop.GetValue(item) ?? DBNull.Value;
                table.Rows.Add(row);
            }

            return table;

        }

        private void dataGridView1_DataBindingComplete(object sender, DataGridViewBindingCompleteEventArgs e)
        {
            switch (btn_sender)
            {
                case "1":
                    #region SAP單
                    DataGridViewCellStyle style = dataGridView1.ColumnHeadersDefaultCellStyle;
                    style.Font = new Font("", 16);
                    dataGridView1.Columns[0].HeaderText = "序";                    
                    dataGridView1.Columns[0].Width = 0;
                    dataGridView1.Columns[0].Visible = false;
                    dataGridView1.Columns[0].SortMode = DataGridViewColumnSortMode.NotSortable;

                    dataGridView1.Columns[1].HeaderText = "製程";
                    dataGridView1.Columns[1].Width = 20;
                    dataGridView1.Columns[1].SortMode = DataGridViewColumnSortMode.NotSortable;

                    dataGridView1.Columns[2].HeaderText = "料號_昶";
                    dataGridView1.Columns[2].Width = 60;
                    dataGridView1.Columns[2].SortMode = DataGridViewColumnSortMode.NotSortable;

                    dataGridView1.Columns[3].HeaderText = "料號_客";
                    dataGridView1.Columns[3].Width = 40;
                    dataGridView1.Columns[3].SortMode = DataGridViewColumnSortMode.NotSortable;

                    dataGridView1.Columns[4].HeaderText = "規格";
                    dataGridView1.Columns[4].Width = 100;
                    dataGridView1.Columns[4].SortMode = DataGridViewColumnSortMode.NotSortable;

                    dataGridView1.Columns[5].HeaderText = "發料倉庫";
                    dataGridView1.Columns[5].Width = 30;
                    dataGridView1.Columns[5].SortMode = DataGridViewColumnSortMode.NotSortable;

                    dataGridView1.Columns[6].HeaderText = "應領數量";
                    dataGridView1.Columns[6].Width = 30;
                    dataGridView1.Columns[6].SortMode = DataGridViewColumnSortMode.NotSortable;

                    dataGridView1.Columns[7].HeaderText = "發料倉庫存";
                    dataGridView1.Columns[7].Width = 35;
                    dataGridView1.Columns[7].SortMode = DataGridViewColumnSortMode.NotSortable;

                    dataGridView1.Columns[8].HeaderText = "KM10庫存數";
                    dataGridView1.Columns[8].Width = 40;
                    dataGridView1.Columns[8].SortMode = DataGridViewColumnSortMode.NotSortable;                    

                    dataGridView1.Columns[9].HeaderText = "替代群";
                    dataGridView1.Columns[9].Width = 30;
                    dataGridView1.Columns[9].SortMode = DataGridViewColumnSortMode.NotSortable;                   

                    dataGridView1.Columns[10].HeaderText = "確認";
                    dataGridView1.Columns[10].Width = 50;
                    dataGridView1.Columns[10].SortMode = DataGridViewColumnSortMode.NotSortable;

                    #endregion
                    break;
                case "2":
                    #region A4單
                    dataGridView1.Columns[0].HeaderText = "序";
                    dataGridView1.Columns[0].Width = 0;
                    dataGridView1.Columns[0].Visible = false;
                    dataGridView1.Columns[0].SortMode = DataGridViewColumnSortMode.NotSortable;

                    dataGridView1.Columns[1].HeaderText = "製程";
                    dataGridView1.Columns[1].Width = 20;
                    dataGridView1.Columns[1].SortMode = DataGridViewColumnSortMode.NotSortable;

                    dataGridView1.Columns[2].HeaderText = "料號_昶";
                    dataGridView1.Columns[2].Width = 60;
                    dataGridView1.Columns[2].SortMode = DataGridViewColumnSortMode.NotSortable;

                    dataGridView1.Columns[3].HeaderText = "料號_客";
                    dataGridView1.Columns[3].Width = 40;
                    dataGridView1.Columns[3].SortMode = DataGridViewColumnSortMode.NotSortable;

                    dataGridView1.Columns[4].HeaderText = "規格";
                    dataGridView1.Columns[4].Width = 120;
                    dataGridView1.Columns[4].SortMode = DataGridViewColumnSortMode.NotSortable;                    

                    dataGridView1.Columns[5].HeaderText = "應領數量";
                    dataGridView1.Columns[5].Width = 40;
                    dataGridView1.Columns[5].SortMode = DataGridViewColumnSortMode.NotSortable;                    

                    dataGridView1.Columns[6].HeaderText = "出庫";
                    dataGridView1.Columns[6].Width = 50;
                    dataGridView1.Columns[6].SortMode = DataGridViewColumnSortMode.NotSortable;
                    #endregion
                    break;
                default:
                    break;
            }
            
        }
        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == 13)
            {
                keypress(null);
            }
        }
        string filePath = string.Empty;
        private void btn_export_Click(object sender, EventArgs e)
        {
            selectButton.Enabled = false;
            btn_export.Enabled = false;
            commitButton.Enabled = false;        
            filePath = "";
            //string strFilePath = txt_path.Text;
            string newName = txt_engsr.Text.ToUpper() + "_" + txt_wono.Text + "_" + DateTime.Now.ToString("yyyyMMdd");
            //檔案路徑
            //\\192.168.4.11\33 倉庫備料單\2020年備料單\備料完成待扣帳\
            filePath = @"\\miss01\33 倉庫備料單\2020年備料單\備料完成待扣帳\" + newName;

            //原始檔案
            //filePath = FileUpload1.FileName;

            #region 程式路徑Upload資料夾+時間戳記
            //string filename_full = Path.GetFileNameWithoutExtension(filePath) + newName + Path.GetExtension(filePath);
            //string _newfilePath = Application.StartupPath + "\\Upload\\" + filename_full;
            //File.Copy(filePath, _newfilePath, true);
            //filePath = _newfilePath;
            #endregion

            //System.IO.FileInfo batchItemAttribute = new FileInfo(filePath)
            //{
            //    //設定檔案屬性為非唯讀
            //    Attributes = FileAttributes.Normal
            //};
            this.expExcelBW.WorkerSupportsCancellation = true; //允許中斷
            this.expExcelBW.RunWorkerAsync(); //呼叫背景程式            
         }
        private void expExcelSheet()
        {
            Form.CheckForIllegalCrossThreadCalls = false;
            try
            {
                progressBar1.Minimum = 0;
                progressBar1.Maximum = table.Rows.Count;
                progressBar1.Step = 1;
                
                //應用程序
                Excel.Application excel_app = new Excel.Application();
                //開啟存在的檔案
                //Excel.Workbook excel_wb = excel_app.Workbooks.Open(filePath);
                //檔案
                Excel.Workbook excel_wb = excel_app.Workbooks.Add();
                //工作表
                Excel.Worksheet excel_ws = new Excel.Worksheet();
                excel_ws = excel_wb.Worksheets[1];                
                excel_ws.Name = "Sheet1";
                int x = (btn_sender == "1") ? 10 : 6; //(OLD)
                //int x = (btn_sender=="1")?18:6; //(NEW)
                for (int i = 0; i <= x; i++)
                {
                    excel_app.Cells[1, i + 1] = dataGridView1.Columns[i].HeaderText;
                }
                for (int i = 0; i < table.Rows.Count; i++)
                {
                    excel_app.Cells.NumberFormat = "@";
                    excel_app.Cells[i + 2, 1] = i + 1;
                    excel_app.Cells[i + 2, 2] = table.Rows[i][1].ToString();
                    excel_app.Cells[i + 2, 3] = table.Rows[i][2].ToString();
                    excel_app.Cells[i + 2, 4] = table.Rows[i][3].ToString();
                    excel_app.Cells[i + 2, 5] = table.Rows[i][4].ToString();
                    excel_app.Cells[i + 2, 6] = table.Rows[i][5].ToString();
                    excel_app.Cells[i + 2, 7] = table.Rows[i][6].ToString();                    
                    if (btn_sender=="1")
                    {
                        excel_app.Cells[i + 2, 8] = table.Rows[i][7].ToString();
                        excel_app.Cells[i + 2, 9] = table.Rows[i][8].ToString();
                        excel_app.Cells[i + 2, 10] = table.Rows[i][9].ToString();
                        excel_app.Cells[i + 2, 11] = table.Rows[i][10].ToString();
                        //excel_app.Cells[i + 2, 12] = table.Rows[i][11].ToString();
                        //excel_app.Cells[i + 2, 13] = table.Rows[i][12].ToString();
                        //excel_app.Cells[i + 2, 14] = table.Rows[i][13].ToString();
                        //excel_app.Cells[i + 2, 15] = table.Rows[i][14].ToString();
                        //excel_app.Cells[i + 2, 16] = table.Rows[i][15].ToString();
                        //excel_app.Cells[i + 2, 17] = table.Rows[i][16].ToString();
                        //excel_app.Cells[i + 2, 18] = table.Rows[i][17].ToString();
                    }
                                        
                    progressBar1.PerformStep();
                }

                //存檔
                excel_wb.SaveAs(filePath, Excel.XlFileFormat.xlExcel8);
                
                //關閉book
                excel_wb.Close(false, Type.Missing, Type.Missing);
                //關閉excel
                excel_app.Quit();
                //關閉&釋放                
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excel_app);
                excel_ws = null;                
                excel_wb = null;
                excel_app = null;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private string subShift(string txt)
        {
            string sub = StringSplit.StrLeft(StringSplit.StrRight(txt, 3), 1);
            string strShift = (sub == "-" && (StringSplit.StrRight(txt, 3).CompareTo("-01") != 0
                                            || StringSplit.StrRight(txt, 3).CompareTo("-02") != 0
                                            || StringSplit.StrRight(txt, 3).CompareTo("-03") != 0
                                            || StringSplit.StrRight(txt, 3).CompareTo("-04") != 0
                                            || StringSplit.StrRight(txt, 3).CompareTo("-05") != 0))
                       ? StringSplit.StrLeft(txt, txt.Length - 3) : txt;

            return strShift;
        }
        DataSet dsRE = new DataSet();
        private void txt_wono_Leave(object sender, EventArgs e)
        {
            string strRE = @"select Item_No_Master ,sum(amount) as amount from Automatic_Storage_Output where Wo_No=@wono group by Item_No_Master ";
            SqlParameter[] parmRE = new SqlParameter[] 
            {
                //new SqlParameter("engsr",txt_engSR.Text.Trim().ToUpper()),
                new SqlParameter("wono",txt_wono.Text.Trim())
            };
            dsRE = db.ExecuteDataSet(strRE, CommandType.Text, parmRE);
        }
        private void loadExcelBW_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            progressBar1.Value = e.ProgressPercentage;
        }

        private void expExcelBW_DoWork(object sender, DoWorkEventArgs e)
        {
            if (expExcelBW.CancellationPending) //如果被中斷...
                e.Cancel = true;
            expExcelSheet();
            GC.Collect();
            this.Close();
        }

        private void expExcelBW_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            MessageBox.Show("匯出完成");
            btn_export.Enabled = true;
            commitButton.Enabled = true;
            selectButton.Enabled = true;
        }

        private void expExcelBW_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            progressBar1.Value = e.ProgressPercentage;
        }
        int VerticalScrollIndex = 0, HorizontalOffset = 0;
        private void dataGridView1_Scroll(object sender, ScrollEventArgs e)
        {
            try
            {
                if (e.ScrollOrientation == ScrollOrientation.VerticalScroll)
                {
                    VerticalScrollIndex = e.NewValue;
                }
                else if (e.ScrollOrientation == ScrollOrientation.HorizontalScroll)
                {
                    HorizontalOffset = e.NewValue;
                }

            }
            catch { }
        }
        private void BatOut_Resize(object sender, EventArgs e)
        {
            if (X == 0 || Y == 0) return;
            fgX = (float)this.Width / (float)X;
            fgY = (float)this.Height / (float)Y;

            SetControls(fgX, fgY, this);
        }

        private void BatOut_Shown(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Maximized;
        }
        /// <summary>
        /// 將控制項的寬，高，左邊距，頂邊距和字體大小暫存到tag屬性中
        /// </summary>
        /// <param name="cons">遞歸控制項中的控制項</param>
        private void SetTag(Control cons)
        {
            foreach (Control con in cons.Controls)
            {
                con.Tag = con.Width + ":" + con.Height + ":" + con.Left + ":" + con.Top + ":" + con.Font.Size;
                if (con.Controls.Count > 0)
                    SetTag(con);
            }
        }

        private void btn_Out_Click(object sender, EventArgs e)
        {
            OutPut output = new OutPut();
            output.Owner = this;
            output.Show();
        }

        private void dataGridView1_RowStateChanged(object sender, DataGridViewRowStateChangedEventArgs e)
        {
            //e.Row.HeaderCell.Value = string.Format("{0}", e.Row.Index + 1);
        }

        private void BatOut_FormClosing(object sender, FormClosingEventArgs e)
        {
            ClearMemory();            
        }

        private void SetControls(float newx, float newy, Control cons)
        {
            if (isLoaded)
            {
                //遍歷窗體中的控制項，重新設置控制項的值
                foreach (Control con in cons.Controls)
                {
                    string[] mytag = con.Tag.ToString().Split(new char[] { ':' });//獲取控制項的Tag屬性值，並分割後存儲字元串數組
                    float a = System.Convert.ToSingle(mytag[0]) * newx;//根據窗體縮放比例確定控制項的值，寬度
                    con.Width = (int)a;//寬度
                    a = System.Convert.ToSingle(mytag[1]) * newy;//高度
                    con.Height = (int)(a);
                    a = System.Convert.ToSingle(mytag[2]) * newx;//左邊距離
                    con.Left = (int)(a);
                    a = System.Convert.ToSingle(mytag[3]) * newy;//上邊緣距離
                    con.Top = (int)(a);
                    Single currentSize = System.Convert.ToSingle(mytag[4]) * newy;//字體大小
                    con.Font = new Font(con.Font.Name, currentSize, con.Font.Style, con.Font.Unit);
                    if (con.Controls.Count > 0)
                    {
                        SetControls(newx, newy, con);
                    }
                }
            }
        }

        private void dataGridView1_DoubleClick(object sender, EventArgs e)
        {
            childValueButton = false;
            OutPut output = new OutPut();
            output.Item = dataGridView1.CurrentRow.Cells[2].Value.ToString();            
            output.Spec = dataGridView1.CurrentRow.Cells[4].Value.ToString();
            output.Wono = txt_wono.Text.Trim().ToUpper();
            output.Bat = true;

            output.setValue();
            output.Owner = this;
            output.ShowDialog();

            //    //子視窗關閉會回來這裡
            //    //回傳位置
            //if (!string.IsNullOrEmpty(childValuePosition) || !string.IsNullOrEmpty(childValueEngSr))
            //{
            //    keypress(childValuePosition);
            //}
            //if (!string.IsNullOrEmpty(childValueEngSr))
            //{
            //    BarOutViewReBack(childValuePosition);
            //}
            if (childValueButton)
            {
                if (!string.IsNullOrEmpty(childValueChk) && !string.IsNullOrEmpty(childValueItemE))
                {
                    ChildFormReturnChk();
                }
            }
            
        }
        public void ChildFormReturnChk()
        {
            string shiftstr = childValueItemE;
            //更新畫面DataGV    
            #region LINQ
            int c = 0, sn = 0, repeat = 0;
            string position = "", engsr = "";
            var rowsToUpdate =
                    table.AsEnumerable().Where(r => r.Field<string>("itemE") == shiftstr);
            foreach (var row in rowsToUpdate)
            {
                c++;
                int count_A = Convert.ToInt32(childValueChk);
                int ck_i = (string.IsNullOrEmpty(row["chk"].ToString())) ? 0 : Convert.ToInt32(row["chk"]);
                int ck = ck_i + count_A;
                row.SetField("chk", ck);
                //row.SetField("chk", "OK");
                //row.SetField("engsr", txt_engSR.Text.Trim().ToUpper());
                sn = Convert.ToInt32(row["sno"].ToString());

                //engsr = txt_engSR.Text.Trim().ToUpper();
            }
            #endregion
            dataGridView1.DataSource = null;
            dataGridView1.DataSource = table;

            dataGridView1.ClearSelection();
            switch (btn_sender)
            {
                case "1"://SAP
                    dataGridView1.Rows[sn - 2].Selected = true;
                    dataGridView1.FirstDisplayedScrollingRowIndex = (sn - 2);
                    break;
                case "2"://A4
                    dataGridView1.Rows[sn - 3].Selected = true;
                    dataGridView1.FirstDisplayedScrollingRowIndex = (sn - 3);
                    break;
                default:
                    break;
            }
            
        }
        #region SMD的扣帳方法
        public void keypress(string outinPosition)
        {
            #region WritetoList
            string shiftstr = txt_item2.Text.Trim().ToUpper();
            #endregion

            //更新畫面DataGV    
            #region LINQ
            int c = 0, sn = 0, repeat = 0;
            string position = "", engsr = "";
            var rowsToUpdate =
                    cpDataTable.AsEnumerable().Where(r => r.Field<string>("item2") == shiftstr);
            foreach (var row in rowsToUpdate)
            {
                c++;
                if (Convert.ToString(row["chk"].ToString()) == "OK")
                {
                    repeat = 2;
                }
                row.SetField("chk", "OK");
                //row.SetField("engsr", txt_engSR.Text.Trim().ToUpper());
                sn = Convert.ToInt32(row["sno"].ToString());
                position = (!string.IsNullOrWhiteSpace(outinPosition)) ? outinPosition : row["position"].ToString();
                //engsr = txt_engSR.Text.Trim().ToUpper();
            }
            #endregion

            //if (obj != null)
            if (c > 0)
            {
                //實際扣帳
                #region 扣帳程序
                try
                {
                    //如果View視窗點工程編號則不進行扣帳程序
                    if (string.IsNullOrEmpty(childValueEngSr))
                    {
                        //Insert Out_Storage
                        string sqlOut = @"insert into Automatic_Storage_Output (Sno,Item_No_Master,Position,Unit_No,Output_UserNo,Output_Date,Wo_No,Eng_SR) 
                                                   select sno,Item_No_Master,Position,@unitno,@outuser,@outdate,@wono,@engsr from Automatic_Storage_Detail 
                                                   where Item_No_Master=@master and Position=@position and amount > 0";
                        SqlParameter[] parm1 = new SqlParameter[]
                        {
                            new SqlParameter("unitno",Login.Unit_No),
                            new SqlParameter("outuser",Login.User_No),
                            new SqlParameter("outdate",OutDate),
                            new SqlParameter("wono",txt_wono.Text.Trim().ToUpper()),
                            //new SqlParameter("engsr",txt_engSR.Text.Trim().ToUpper()),
                            new SqlParameter("master", txt_item2.Text.Trim().ToUpper()),
                            new SqlParameter("position",position.ToUpper())
                        };
                        db.ExecueNonQuery(sqlOut, CommandType.Text, parm1);
                    }

                    string sqlDetail = "update Automatic_Storage_Detail set ";
                    if (string.IsNullOrEmpty(childValueEngSr))
                    {
                        sqlDetail += "Amount = Amount-1 ,";
                    }
                    sqlDetail += "Up_OutDate = @outdate , Output_UserNo = @outuser ,Eng_SR = @engsr where Item_No_Master= @master and Position=@position";
                    SqlParameter[] parm2 = new SqlParameter[]
                    {
                        new SqlParameter("outdate",OutDate),
                        new SqlParameter("outuser",Login.User_No),
                        //new SqlParameter("engsr",txt_engSR.Text.Trim().ToUpper()),
                        new SqlParameter("master",txt_item2.Text.Trim().ToUpper()),
                        new SqlParameter("position",position.Trim()),
                    };
                    db.ExecueNonQuery(sqlDetail, CommandType.Text, parm2);

                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
                #endregion

                dataGridView1.DataSource = null;
                dataGridView1.DataSource = cpDataTable;

                txt_item2.Text = "";
            }
            else
            {
                if (MessageBox.Show(txt_item2.Text.Trim() + Environment.NewLine + Environment.NewLine + "請確認料卷是否正確", "錯誤!!") == DialogResult.OK)
                {
                    txt_item2.Text = "";
                }
            }

            //if (obj != null)
            if (c > 0)
            {
                //this.dataGridView1.FirstDisplayedCell = this.dataGridView1.CurrentCell;                    
                dataGridView1.ClearSelection();
                dataGridView1.Rows[sn - 1].Selected = true;
                if (repeat == 2)
                {
                    dataGridView1.Rows[sn - 1].DefaultCellStyle.SelectionForeColor = Color.Red;
                    dataGridView1.Rows[sn - 1].DefaultCellStyle.BackColor = Color.LightGreen;
                }
                dataGridView1.FirstDisplayedScrollingRowIndex = (sn - 1);
            }
            txt_item2.Focus();
        }
        #endregion

        #region 子視窗點工程編號
        public void BarOutViewReBack(string point)
        {
            #region WritetoList
            string shiftstr = txt_item2.Text.Trim().ToUpper();
            //var obj = ls2.FirstOrDefault(x => x.Item2.Trim().ToUpper() == shiftstr);
            #endregion

            //更新畫面DataGV    
            #region LINQ
            int c = 0, sn = 0;
            string position = "", engsr = "";
            var rowsToUpdate =
                    cpDataTable.AsEnumerable().Where(r => r.Field<string>("item2") == shiftstr);
            foreach (var row in rowsToUpdate)
            {
                c++;
                row.SetField("chk", "OK");
                //row.SetField("engsr", txt_engSR.Text.Trim().ToUpper());
                sn = Convert.ToInt32(row["sno"].ToString());
                
                //engsr = txt_engSR.Text.Trim().ToUpper();
            }
            #endregion

            string sqlDetail = @"update Automatic_Storage_Detail 
                                                           set Amount = Amount-1 ,Up_OutDate = @outdate , Output_UserNo = @outuser 
                                                           where Item_No_Master= @master and Position=@position";
            SqlParameter[] parm2 = new SqlParameter[]
            {
                        new SqlParameter("outdate",OutDate),
                        new SqlParameter("outuser",Login.User_No),
                        //new SqlParameter("engsr",txt_engSR.Text.Trim().ToUpper()),
                        new SqlParameter("master",txt_item2.Text.Trim().ToUpper()),
                        new SqlParameter("position",position.Trim()),
            };
            db.ExecueNonQuery(sqlDetail, CommandType.Text, parm2);

    

            dataGridView1.DataSource = null;
            dataGridView1.DataSource = cpDataTable;

            dataGridView1.ClearSelection();
            dataGridView1.Rows[sn - 1].Selected = true;
            dataGridView1.FirstDisplayedScrollingRowIndex = (sn - 1);
        }
        #endregion

        #region 記憶體回收
        [DllImport("kernel32.dll", EntryPoint = "SetProcessWorkingSetSize")]
        public static extern int SetProcessWorkingSetSize(IntPtr process, int minSize, int maxSize);

        private void loadExcelBW_DoWork(object sender, DoWorkEventArgs e)
        {

        }

        private void btn_change_Click(object sender, EventArgs e)
        {
            string newPosition = txt_newPosition.Text.ToUpper();
            string sourcePosition = txt_source.Text.ToUpper();

            string[] sqlCommands = {
    "UPDATE Automatic_Storage_Detail SET position = @Position WHERE amount > 0 AND position = @SourcePosition",
    "UPDATE Automatic_Storage_Input SET position = @Position WHERE amount > 0 AND position = @SourcePosition",
    "UPDATE Automatic_Storage_Output SET position = @Position WHERE amount > 0 AND position = @SourcePosition"
};
            try
            {

                string str_s = @"select * from Automatic_Storage_Position where Unit_No=@unitno and Position=@positionS";
                SqlParameter[] parm_s = new SqlParameter[]
                {
                                new SqlParameter("unitno",Login.Unit_No),
                                new SqlParameter("positionS",newPosition),
                };
                DataSet dataSet = db.ExecuteDataSet(str_s, CommandType.Text, parm_s);
                if (dataSet.Tables[0].Rows.Count == 0)
                {
                    string strsql = @"insert into Automatic_Storage_Position (Unit_No,Position,Create_User,Create_Date) values
                                                        (@unitno, @positionA, @cr_user, @cr_date)";
                    SqlParameter[] parameters = new SqlParameter[]
                    {
                                    new SqlParameter("unitno",Login.Unit_No),
                                    new SqlParameter("positionA",newPosition),
                                    new SqlParameter("cr_user",Login.User_No),
                                    new SqlParameter("cr_date",DateTime.Now)
                    };
                    db.ExecueNonQuery(strsql, CommandType.Text, parameters);
                    txt_errLog.Text = "儲位新增完成";
                }

                foreach (string sql in sqlCommands)
                {
                    // Create parameters for each execution
                    SqlParameter[] parameters = 
                    {
                        new SqlParameter("@Position", newPosition),
                        new SqlParameter("@SourcePosition", sourcePosition)
                    };

                    // Execute the query
                    db.ExecueNonQuery(sql, CommandType.Text, parameters);

                    // Optional: Handle the number of affected rows if needed
                    //Console.WriteLine($"Rows affected: {rowsAffected}");
                }
            }
            catch (Exception ex)
            {
                // Log the error or display an error message to the user
                Console.WriteLine("Error: " + ex.Message);
            }
        }

        /// <summary>
        /// 釋放記憶體
        /// </summary>
        public static void ClearMemory()
        {
            GC.Collect();
            GC.WaitForPendingFinalizers();
            if (Environment.OSVersion.Platform == PlatformID.Win32NT)
            {
                SetProcessWorkingSetSize(System.Diagnostics.Process.GetCurrentProcess().Handle, -1, -1);
            }
        }
        #endregion
    }
}
