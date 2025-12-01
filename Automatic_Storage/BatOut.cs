using Automatic_Storage.Utilities;
using System; // 引用 System 命名空間
using System.Collections.Generic; // 引用泛型集合
using System.ComponentModel; // 引用元件模型
using System.Data; // 引用資料處理
using System.Data.OleDb; // 引用 OleDb 資料庫
using System.Data.SqlClient; // 引用 SQL Server 資料庫
using System.Drawing; // 引用繪圖
using System.IO; // 引用檔案處理
using System.Linq; // 引用 LINQ 查詢
using System.Security; // 引用安全性
using System.Text; // 引用字串處理
using System.Windows.Forms; // 引用視窗表單
using Excel = Microsoft.Office.Interop.Excel; // 引用 Excel 互操作
using System.Windows.Threading; // 引用執行緒
using System.Runtime.InteropServices; // 引用互操作服務

namespace Automatic_Storage // 命名空間：Automatic_Storage
{
    /// <summary>
    /// 批次出庫主視窗
    /// </summary>
    public partial class BatOut : Form // BatOut 類別，繼承 Form
    {
        private OpenFileDialog FileUpload1; // 檔案上傳對話框
                                            //private Button selectButton; // 選擇按鈕（未使用）
        DataSet dsData = new DataSet(); // 資料集
        /// <summary>
        /// 主資料表
        /// </summary>
        public DataTable table; // 公開資料表
        DataTable cpDataTable = new DataTable(); // SMD扣帳用資料表
        string OutDate = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"); // 出庫日期
        StringBuilder sb = new StringBuilder(); // 字串組合器
        bool txtFlag = true; // 工單號碼輸入狀態
                             //1→ 批次1 ; 2→批次2 // btn_sender 用於判斷批次類型
        private static string btn_sender = string.Empty; // 按鈕來源
        #region 視窗ReSize
        int X = new int();  // 窗口寬度
        int Y = new int(); // 窗口高度
        float fgX = new float(); // 寬度縮放比例
        float fgY = new float(); // 高度縮放比例
        bool isLoaded;  // 是否已設定各控制的尺寸資料到Tag屬性
        #endregion

        #region 子視窗回傳值
        private bool childValueButton; // 子視窗回傳按鈕狀態
        /// <summary>
        /// 子視窗回傳按鈕狀態
        /// </summary>
        public bool MsgFromChildButton
        {
            set { childValueButton = value; }
        }
        private string childValuePosition; // 子視窗回傳儲位
        /// <summary>
        /// 子視窗回傳儲位
        /// </summary>
        public string MsgFromChildPosition
        {
            set { childValuePosition = value; }
        }
        private string childValueEngSr; // 子視窗回傳工程編號
        /// <summary>
        /// 子視窗回傳工程編號
        /// </summary>
        public string MsgFromChildEngSr
        {
            set { childValueEngSr = value; }
        }
        //↓↓原本chk="OK" 變數不異動直接使用變成Amount領用數量
        /// <summary>
        /// 子視窗回傳領用數量
        /// </summary>
        public string MsgFromChildChk
        {
            set { childValueChk = value; }
        }

        private string childValueChk; // 子視窗回傳領用數量

        private string childValueItemE;
        /// <summary>
        /// 子視窗回傳料號
        /// </summary>
        public string MsgFromChildItemE
        {
            set { childValueItemE = value; }
        }
        private string childValueAmount; // 子視窗回傳數量
        /// <summary>
        /// 子視窗回傳數量
        /// </summary>
        public string MsgFromChildAmount
        {
            set { childValueAmount = value; }
        }

        #endregion


        /// <summary>
        /// 批次出庫項目類別
        /// </summary>
        public class BatItem : IComparable<BatItem>, IEquatable<BatItem>
        {
            /// <summary>
            /// 序號
            /// </summary>
            public int sno { get; set; } // 序號
            /// <summary>
            /// 料號1
            /// </summary>
            public string Item1 { get; set; } // 料號1
            /// <summary>
            /// 製程站點
            /// </summary>
            public string Station { get; set; } // 製程站點
            /// <summary>
            /// 規格
            /// </summary>
            public string Spec { get; set; } // 規格
            /// <summary>
            /// 料號2
            /// </summary>
            public string Item2 { get; set; } // 料號2
            /// <summary>
            /// NA 欄位
            /// </summary>
            public string NA { get; set; } // NA 欄位
            /// <summary>
            /// 確認欄位
            /// </summary>
            public string Chk { get; set; } // 確認欄位
            /// <summary>
            /// 儲位
            /// </summary>
            public string Position { get; set; } // 儲位
            /// <summary>
            /// 工程編號
            /// </summary>
            public string Engsr { get; set; } // 工程編號
            /// <summary>
            /// 備註
            /// </summary>
            public string Memo { get; set; } // 備註

            /// <summary>
            /// 轉為字串
            /// </summary>
            /// <returns>批次項目資訊字串</returns>
            public override string ToString()
            {
                return "sno: " + sno + "   Item1: " + Item1 + "Station: " + Station + "   Spec: " + Spec + "Item2: " + Item2 + "   NA: " + NA +
                    "Chk: " + Chk + "   Position: " + Position + "Engsr:" + Engsr + "Memo: " + Memo;
            }
            /// <summary>
            /// 比較物件是否相等
            /// </summary>
            /// <param name="obj">比較物件</param>
            /// <returns>是否相等</returns>
            public override bool Equals(object obj)
            {
                if (obj == null) return false;
                BatItem objAsPart = obj as BatItem;
                if (objAsPart == null) return false;
                else return Equals(objAsPart);
            }
            /// <summary>
            /// 名稱升冪排序
            /// </summary>
            /// <param name="name1">名稱1</param>
            /// <param name="name2">名稱2</param>
            /// <returns>比較結果</returns>
            public int SortByNameAscending(string name1, string name2)
            {
                return name1.CompareTo(name2);
            }

            // Default comparer for Part type.
            /// <summary>
            /// 比較儲位排序
            /// </summary>
            /// <param name="comparePart">比較物件</param>
            /// <returns>排序結果</returns>
            public int CompareTo(BatItem comparePart)
            {
                // A null value means that this object is greater.
                if (comparePart == null)
                    return 1;

                else
                    return this.Position.CompareTo(comparePart.Position);
            }
            /// <summary>
            /// 取得雜湊碼
            /// </summary>
            /// <returns>雜湊碼</returns>
            public override int GetHashCode()
            {
                return sno;
            }
            /// <summary>
            /// 比較批次項目是否相等
            /// </summary>
            /// <param name="other">另一個批次項目</param>
            /// <returns>是否相等</returns>
            public bool Equals(BatItem other)
            {
                if (other == null) return false;
                return (this.sno.Equals(other.sno));
            }
            // Should also override == and != operators.
        }
        /// <summary>
        /// 建立資料表欄位
        /// </summary>
        /// <returns>資料表</returns>
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
        /// <summary>
        /// 建構函式
        /// </summary>
        public BatOut()
        {
            InitializeComponent();
            isLoaded = false;

        }
        /// <summary>
        /// 視窗載入事件
        /// </summary>
        /// <param name="sender">事件來源</param>
        /// <param name="e">事件參數</param>
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

        /// <summary>
        /// 選擇檔案按鈕事件
        /// </summary>
        /// <param name="sender">事件來源</param>
        /// <param name="e">事件參數</param>
        private void selectButton_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(txt_wono.Text))
            {
                txtFlag = false;
                MessageBox.Show("請先輸入工單號碼");
                return;
            }
            if (txtFlag)
            {
                if (FileUpload1.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        var fileName = FileUpload1.FileName;

                        FileInfo _info = new FileInfo(fileName);
                        string uploadPath = UploadHelper.GetUploadFilePath(_info.Name);
                        if (File.Exists(fileName))
                        {
                            try
                            {
                                _info.CopyTo(uploadPath, true);
                                txt_path.Text = uploadPath;
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show($"複製檔案失敗: {ex.Message}");
                            }
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

        /// <summary>
        /// 批次1提交按鈕事件
        /// </summary>
        /// <param name="sender">事件來源</param>
        /// <param name="e">事件參數</param>
        private void commitButton_Click(object sender, EventArgs e)
        {
            //this.loadExcelBW.WorkerSupportsCancellation = true; //允許中斷
            //this.loadExcelBW.RunWorkerAsync(); //呼叫背景程式
            btn_sender = "1";
            table = table ?? columnsData();
            LoadExcel();
        }
        /// <summary>
        /// 批次2提交按鈕事件
        /// </summary>
        /// <param name="sender">事件來源</param>
        /// <param name="e">事件參數</param>
        private void commitButton2_Click(object sender, EventArgs e)
        {
            btn_sender = "2";
            table = table ?? columnsData();
            LoadExcel();
        }
        /// <summary>
        /// 載入 Excel 資料
        /// </summary>
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
                        progressBar1.Maximum = dsData.Tables[0].Rows.Count - 1;
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
        /// <summary>
        /// 泛型集合轉換為 DataTable
        /// </summary>
        /// <typeparam name="T">資料型別</typeparam>
        /// <param name="data">資料集合</param>
        /// <returns>DataTable</returns>
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

        /// <summary>
        /// DataGridView 資料繫結完成事件
        /// </summary>
        /// <param name="sender">事件來源</param>
        /// <param name="e">事件參數</param>
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

        /// <summary>
        /// 檔案路徑
        /// </summary>
        string filePath = string.Empty;

        /// <summary>
        /// 匯出 Excel 按鈕事件
        /// </summary>
        /// <param name="sender">事件來源</param>
        /// <param name="e">事件參數</param>
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

        /// <summary>
        /// 匯出 Excel 工作表
        /// </summary>
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
                Excel.Worksheet excel_ws = (Excel.Worksheet)excel_wb.Worksheets[1];
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
                    if (btn_sender == "1")
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
                // 提前宣告在 try 內會被建立的 COM 物件，讓外層可用於最後釋放
                Excel.Range colRange = null;
                Excel.Worksheet recordWs = null;
                Excel.Worksheet mainWs = null;

                // 在儲存前，確保工作表保護與應領數量欄位被標記為 Locked
                try
                {
                    string protectPassword = "1234";
                    // 確保主工作表存在
                    mainWs = (Excel.Worksheet)excel_wb.Worksheets[1];
                    // 找到應領數量欄位的索引（比對標題）
                    int amountColIndex = -1;
                    int colCount = mainWs.UsedRange.Columns.Count;
                    for (int c = 1; c <= colCount; c++)
                    {
                        var header = (mainWs.Cells[1, c] as Excel.Range)?.Value2;
                        if (header != null && header.ToString() == "應領數量")
                        {
                            amountColIndex = c;
                            break;
                        }
                    }

                    // 將該欄標為 Locked（若找得到）
                    if (amountColIndex > 0)
                    {
                        colRange = (Excel.Range)mainWs.Columns[amountColIndex];
                        colRange.Locked = true;
                    }

                    // 確保 '記錄' 工作表存在；若不存在則建立並放在第二個位置
                    // recordWs 已在外層宣告，這裡直接使用
                    bool foundRecord = false;
                    for (int i = 1; i <= excel_wb.Worksheets.Count; i++)
                    {
                        var ws = excel_wb.Worksheets[i] as Excel.Worksheet;
                        if (ws != null && ws.Name == "記錄")
                        {
                            recordWs = ws;
                            foundRecord = true;
                            break;
                        }
                    }
                    if (!foundRecord)
                    {
                        try { recordWs = excel_wb.Worksheets.Add(After: excel_wb.Worksheets[1]) as Excel.Worksheet; } catch { recordWs = excel_wb.Worksheets.Add(Type.Missing, excel_wb.Worksheets[1], Type.Missing, Type.Missing) as Excel.Worksheet; }
                        recordWs.Name = "記錄";
                    }

                    // Protect the sheets with password
                    try
                    {
                        // 保護主工作表
                        mainWs.Protect(protectPassword, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                    }
                    catch { }
                    try
                    {
                        // 保護記錄工作表
                        recordWs.Protect(protectPassword, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                    }
                    catch { }
                }
                catch { }

                // 釋放在此區段建立的暫時 COM 物件，避免殘留
                try { if (colRange != null) Automatic_Storage.Utilities.ComInterop.ReleaseComObjectSafe(colRange); } catch { }
                try { if (recordWs != null) Automatic_Storage.Utilities.ComInterop.ReleaseComObjectSafe(recordWs); } catch { }
                try { if (mainWs != null) Automatic_Storage.Utilities.ComInterop.ReleaseComObjectSafe(mainWs); } catch { }

                excel_wb.SaveAs(filePath, Excel.XlFileFormat.xlExcel8);

                //關閉book
                excel_wb.Close(false, Type.Missing, Type.Missing);
                //關閉excel
                excel_app.Quit();
                //關閉&釋放：先釋放 workbook/worksheet，再釋放 application
                try { if (excel_wb != null) Automatic_Storage.Utilities.ComInterop.ReleaseComObjectSafe(excel_wb); } catch { }
                try { if (excel_ws != null) Automatic_Storage.Utilities.ComInterop.ReleaseComObjectSafe(excel_ws); } catch { }
                try { Automatic_Storage.Utilities.ComInterop.ReleaseComObjectSafe(excel_app); } catch { }
                excel_ws = null;
                excel_wb = null;
                excel_app = null;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        /// <summary>
        /// 取得料卷字串
        /// </summary>
        /// <param name="txt">原始字串</param>
        /// <returns>處理後字串</returns>
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

        /// <summary>
        /// 出庫紀錄資料集
        /// </summary>
        /// <remarks>
        /// 用於儲存工單出庫相關的查詢結果。
        /// </remarks>
        /// <example>
        /// dsRE = db.ExecuteDataSet(strRE, CommandType.Text, parmRE);
        /// </example>
        DataSet dsRE = new DataSet(); // 出庫紀錄資料集

        /// <summary>
        /// 工單號碼離開事件
        /// </summary>
        /// <param name="sender">事件來源</param>
        /// <param name="e">事件參數</param>
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

        /// <summary>
        /// 背景工作進度改變事件
        /// </summary>
        /// <param name="sender">事件來源</param>
        /// <param name="e">事件參數</param>
        private void loadExcelBW_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            progressBar1.Value = e.ProgressPercentage;
        }

        /// <summary>
        /// 匯出 Excel 背景工作執行事件
        /// </summary>
        /// <param name="sender">事件來源</param>
        /// <param name="e">事件參數</param>
        private void expExcelBW_DoWork(object sender, DoWorkEventArgs e)
        {
            if (expExcelBW.CancellationPending) //如果被中斷...
                e.Cancel = true;
            expExcelSheet();
            GC.Collect();
            this.Close();
        }

        /// <summary>
        /// 匯出 Excel 背景工作完成事件
        /// </summary>
        /// <param name="sender">事件來源</param>
        /// <param name="e">事件參數</param>
        private void expExcelBW_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            MessageBox.Show("匯出完成");
            btn_export.Enabled = true;
            commitButton.Enabled = true;
            selectButton.Enabled = true;
        }

        /// <summary>
        /// 匯出 Excel 背景工作進度改變事件
        /// </summary>
        /// <param name="sender">事件來源</param>
        /// <param name="e">事件參數</param>
        private void expExcelBW_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            progressBar1.Value = e.ProgressPercentage;
        }
        int VerticalScrollIndex = 0, HorizontalOffset = 0;

        /// <summary>
        /// DataGridView 捲動事件
        /// </summary>
        /// <param name="sender">事件來源</param>
        /// <param name="e">事件參數</param>
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

        /// <summary>
        /// 視窗大小調整事件
        /// </summary>
        /// <param name="sender">事件來源</param>
        /// <param name="e">事件參數</param>
        private void BatOut_Resize(object sender, EventArgs e)
        {
            if (X == 0 || Y == 0) return;
            fgX = (float)this.Width / (float)X;
            fgY = (float)this.Height / (float)Y;

            SetControls(fgX, fgY, this);
        }

        /// <summary>
        /// 視窗顯示事件
        /// </summary>
        /// <param name="sender">事件來源</param>
        /// <param name="e">事件參數</param>
        private void BatOut_Shown(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Maximized;
        }

        /// <summary>
        /// 將控制項的寬，高，左邊距，頂邊距和字體大小暫存到tag屬性中
        /// </summary>
        /// <param name="cons">遞歸控制項中的控制項</param>
        /// <summary>
        /// 設定控制項 Tag 屬性
        /// </summary>
        /// <param name="cons">控制項</param>
        private void SetTag(Control cons)
        {
            foreach (Control con in cons.Controls)
            {
                con.Tag = con.Width + ":" + con.Height + ":" + con.Left + ":" + con.Top + ":" + con.Font.Size;
                if (con.Controls.Count > 0)
                    SetTag(con);
            }
        }

        /// <summary>
        /// 出庫按鈕事件
        /// </summary>
        /// <param name="sender">事件來源</param>
        /// <param name="e">事件參數</param>
        private void btn_Out_Click(object sender, EventArgs e)
        {
            OutPut output = new OutPut();
            output.Owner = this;
            output.Show();
        }

        /// <summary>
        /// DataGridView 行狀態改變事件
        /// </summary>
        /// <param name="sender">事件來源</param>
        /// <param name="e">事件參數</param>
        private void dataGridView1_RowStateChanged(object sender, DataGridViewRowStateChangedEventArgs e)
        {
            //e.Row.HeaderCell.Value = string.Format("{0}", e.Row.Index + 1);
        }

        /// <summary>
        /// 視窗關閉事件
        /// </summary>
        /// <param name="sender">事件來源</param>
        /// <param name="e">事件參數</param>
        private void BatOut_FormClosing(object sender, FormClosingEventArgs e)
        {
            ClearMemory();
        }

        /// <summary>
        /// 設定控制項尺寸
        /// </summary>
        /// <param name="newx">寬度比例</param>
        /// <param name="newy">高度比例</param>
        /// <param name="cons">控制項</param>
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

        /// <summary>
        /// DataGridView 雙擊事件
        /// </summary>
        /// <param name="sender">事件來源</param>
        /// <param name="e">事件參數</param>
        private void dataGridView1_DoubleClick(object sender, EventArgs e)
        {
            childValueButton = false;
            OutPut output = new OutPut();
            output.Item = dataGridView1.CurrentRow?.Cells[2].Value?.ToString() ?? string.Empty;
            output.Spec = dataGridView1.CurrentRow?.Cells[4].Value?.ToString() ?? string.Empty;
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

        /// <summary>
        /// 子視窗回傳確認
        /// </summary>
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
        /// <summary>
        /// SMD 扣帳方法
        /// </summary>
        /// <param name="outinPosition">儲位</param>
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
            // 將焦點移至儲位輸入框
            txt_item2.Focus();
        }
        #endregion

        #region 子視窗點工程編號
        /// <summary>
        /// 子視窗點工程編號回傳
        /// </summary>
        /// <param name="point">工程編號</param>
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
                // 使用集中式 Logger 非同步寫入錯誤，避免直接輸出到 Console
                try
                {
                    _ = System.Threading.Tasks.Task.Run(() => Automatic_Storage.Utilities.Logger.LogErrorAsync("BatOut Error: " + ex.Message));
                }
                catch { }
            }
        }

        /// <summary>
        /// 釋放記憶體
        /// </summary>
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
