
namespace Automatic_Storage
{
    partial class Form1
    {
        /// <summary>
        /// 設計工具所需的變數。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 清除任何使用中的資源。
        /// </summary>
        /// <param name="disposing">如果應該處置受控資源則為 true，否則為 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form 設計工具產生的程式碼

        /// <summary>
        /// 此為設計工具支援所需的方法 - 請勿使用程式碼編輯器修改
        /// 這個方法的內容。
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.btn_Input = new System.Windows.Forms.Button();
            this.btn_Out = new System.Windows.Forms.Button();
            this.btn_BatIn = new System.Windows.Forms.Button();
            this.btn_BatOut = new System.Windows.Forms.Button();
            this.panel1 = new System.Windows.Forms.Panel();
            this.txt_result = new System.Windows.Forms.TextBox();
            this.selectButton = new System.Windows.Forms.Button();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.list_result = new System.Windows.Forms.ListBox();
            this.btn_return = new System.Windows.Forms.Button();
            this.commitButton = new System.Windows.Forms.Button();
            this.txt_path = new System.Windows.Forms.TextBox();
            this.panel2 = new System.Windows.Forms.Panel();
            this.rad_out = new System.Windows.Forms.RadioButton();
            this.rad_in = new System.Windows.Forms.RadioButton();
            this.txt_DTe = new System.Windows.Forms.TextBox();
            this.label9 = new System.Windows.Forms.Label();
            this.txt_DTs = new System.Windows.Forms.TextBox();
            this.label6 = new System.Windows.Forms.Label();
            this.txt_spec = new System.Windows.Forms.TextBox();
            this.btn_qexport = new System.Windows.Forms.Button();
            this.btn_delPosition = new System.Windows.Forms.Button();
            this.btn_combi = new System.Windows.Forms.Button();
            this.button1 = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.txt_position = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.txt_item = new System.Windows.Forms.TextBox();
            this.panel3 = new System.Windows.Forms.Panel();
            this.panel4 = new System.Windows.Forms.Panel();
            this.btn_history = new System.Windows.Forms.Button();
            this.btn_Maintain = new System.Windows.Forms.Button();
            this.fileSystemWatcher1 = new System.IO.FileSystemWatcher();
            this.inputExcelBW = new System.ComponentModel.BackgroundWorker();
            this.panel5 = new System.Windows.Forms.Panel();
            this.label8 = new System.Windows.Forms.Label();
            this.txt_wonoOut = new System.Windows.Forms.TextBox();
            this.label7 = new System.Windows.Forms.Label();
            this.txt_specP5 = new System.Windows.Forms.TextBox();
            this.btn_reP2 = new System.Windows.Forms.Button();
            this.btn_itemSite = new System.Windows.Forms.Button();
            this.btn_findAll = new System.Windows.Forms.Button();
            this.label3 = new System.Windows.Forms.Label();
            this.txt_siteP5 = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.txt_itemP5 = new System.Windows.Forms.TextBox();
            this.outputExcelBW = new System.ComponentModel.BackgroundWorker();
            this.progressBar2 = new System.Windows.Forms.ProgressBar();
            this.panel6 = new System.Windows.Forms.Panel();
            this.txt_sumC = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.rad_E = new System.Windows.Forms.RadioButton();
            this.rad_C = new System.Windows.Forms.RadioButton();
            this.panel_rad = new System.Windows.Forms.Panel();
            this.btnNextPage = new System.Windows.Forms.Button();
            this.btnPreviousPage = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.panel1.SuspendLayout();
            this.panel2.SuspendLayout();
            this.panel3.SuspendLayout();
            this.panel4.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.fileSystemWatcher1)).BeginInit();
            this.panel5.SuspendLayout();
            this.panel6.SuspendLayout();
            this.panel_rad.SuspendLayout();
            this.SuspendLayout();
            // 
            // dataGridView1
            // 
            this.dataGridView1.AllowUserToAddRows = false;
            this.dataGridView1.AllowUserToDeleteRows = false;
            this.dataGridView1.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this.dataGridView1.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.DisplayedCells;
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Location = new System.Drawing.Point(12, 132);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.ReadOnly = true;
            this.dataGridView1.RowTemplate.Height = 24;
            this.dataGridView1.Size = new System.Drawing.Size(776, 351);
            this.dataGridView1.TabIndex = 0;
            this.dataGridView1.CellFormatting += new System.Windows.Forms.DataGridViewCellFormattingEventHandler(this.dataGridView1_CellFormatting);
            this.dataGridView1.DataBindingComplete += new System.Windows.Forms.DataGridViewBindingCompleteEventHandler(this.dataGridView1_DataBindingComplete);
            this.dataGridView1.DoubleClick += new System.EventHandler(this.dataGridView1_DoubleClick);
            // 
            // btn_Input
            // 
            this.btn_Input.BackColor = System.Drawing.SystemColors.ActiveCaption;
            this.btn_Input.Font = new System.Drawing.Font("微軟正黑體", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.btn_Input.Location = new System.Drawing.Point(3, 7);
            this.btn_Input.Name = "btn_Input";
            this.btn_Input.Size = new System.Drawing.Size(89, 37);
            this.btn_Input.TabIndex = 1;
            this.btn_Input.Text = "入庫登記";
            this.btn_Input.UseVisualStyleBackColor = false;
            this.btn_Input.Click += new System.EventHandler(this.button1_Click);
            // 
            // btn_Out
            // 
            this.btn_Out.BackColor = System.Drawing.Color.MistyRose;
            this.btn_Out.Font = new System.Drawing.Font("微軟正黑體", 12F, System.Drawing.FontStyle.Bold);
            this.btn_Out.Location = new System.Drawing.Point(225, 7);
            this.btn_Out.Name = "btn_Out";
            this.btn_Out.Size = new System.Drawing.Size(87, 37);
            this.btn_Out.TabIndex = 2;
            this.btn_Out.Text = "出庫登記";
            this.btn_Out.UseVisualStyleBackColor = false;
            this.btn_Out.Click += new System.EventHandler(this.button2_Click);
            // 
            // btn_BatIn
            // 
            this.btn_BatIn.BackColor = System.Drawing.Color.Gold;
            this.btn_BatIn.Font = new System.Drawing.Font("微軟正黑體", 12F, System.Drawing.FontStyle.Bold);
            this.btn_BatIn.Location = new System.Drawing.Point(454, 7);
            this.btn_BatIn.Name = "btn_BatIn";
            this.btn_BatIn.Size = new System.Drawing.Size(89, 37);
            this.btn_BatIn.TabIndex = 3;
            this.btn_BatIn.Text = "批次入庫";
            this.btn_BatIn.UseVisualStyleBackColor = false;
            this.btn_BatIn.Click += new System.EventHandler(this.button3_Click);
            // 
            // btn_BatOut
            // 
            this.btn_BatOut.BackColor = System.Drawing.Color.Teal;
            this.btn_BatOut.Font = new System.Drawing.Font("微軟正黑體", 12F, System.Drawing.FontStyle.Bold);
            this.btn_BatOut.Location = new System.Drawing.Point(681, 7);
            this.btn_BatOut.Name = "btn_BatOut";
            this.btn_BatOut.Size = new System.Drawing.Size(89, 37);
            this.btn_BatOut.TabIndex = 4;
            this.btn_BatOut.Text = "批次出庫";
            this.btn_BatOut.UseVisualStyleBackColor = false;
            this.btn_BatOut.Click += new System.EventHandler(this.button4_Click);
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.txt_result);
            this.panel1.Controls.Add(this.selectButton);
            this.panel1.Controls.Add(this.progressBar1);
            this.panel1.Controls.Add(this.list_result);
            this.panel1.Controls.Add(this.btn_return);
            this.panel1.Controls.Add(this.commitButton);
            this.panel1.Controls.Add(this.txt_path);
            this.panel1.Location = new System.Drawing.Point(18, 57);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(677, 70);
            this.panel1.TabIndex = 5;
            // 
            // txt_result
            // 
            this.txt_result.Location = new System.Drawing.Point(554, 2);
            this.txt_result.Multiline = true;
            this.txt_result.Name = "txt_result";
            this.txt_result.ScrollBars = System.Windows.Forms.ScrollBars.Horizontal;
            this.txt_result.Size = new System.Drawing.Size(120, 64);
            this.txt_result.TabIndex = 12;
            // 
            // selectButton
            // 
            this.selectButton.Font = new System.Drawing.Font("微軟正黑體", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.selectButton.Location = new System.Drawing.Point(293, 17);
            this.selectButton.Name = "selectButton";
            this.selectButton.Size = new System.Drawing.Size(87, 37);
            this.selectButton.TabIndex = 11;
            this.selectButton.Text = "選擇檔案";
            this.selectButton.UseVisualStyleBackColor = true;
            this.selectButton.Click += new System.EventHandler(this.selectButton_Click);
            // 
            // progressBar1
            // 
            this.progressBar1.Location = new System.Drawing.Point(14, 53);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(244, 13);
            this.progressBar1.TabIndex = 9;
            // 
            // list_result
            // 
            this.list_result.FormattingEnabled = true;
            this.list_result.ItemHeight = 12;
            this.list_result.Location = new System.Drawing.Point(554, 2);
            this.list_result.Name = "list_result";
            this.list_result.Size = new System.Drawing.Size(120, 64);
            this.list_result.TabIndex = 10;
            // 
            // btn_return
            // 
            this.btn_return.Location = new System.Drawing.Point(492, 30);
            this.btn_return.Name = "btn_return";
            this.btn_return.Size = new System.Drawing.Size(50, 23);
            this.btn_return.TabIndex = 9;
            this.btn_return.Text = "返回";
            this.btn_return.UseVisualStyleBackColor = true;
            this.btn_return.Click += new System.EventHandler(this.btn_return_Click);
            // 
            // commitButton
            // 
            this.commitButton.Font = new System.Drawing.Font("微軟正黑體", 12F, System.Drawing.FontStyle.Bold);
            this.commitButton.Location = new System.Drawing.Point(397, 17);
            this.commitButton.Name = "commitButton";
            this.commitButton.Size = new System.Drawing.Size(89, 37);
            this.commitButton.TabIndex = 8;
            this.commitButton.Text = "上傳入庫";
            this.commitButton.UseVisualStyleBackColor = true;
            this.commitButton.Click += new System.EventHandler(this.commitButton_Click);
            // 
            // txt_path
            // 
            this.txt_path.Location = new System.Drawing.Point(14, 25);
            this.txt_path.Name = "txt_path";
            this.txt_path.Size = new System.Drawing.Size(244, 22);
            this.txt_path.TabIndex = 6;
            // 
            // panel2
            // 
            this.panel2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel2.Controls.Add(this.rad_out);
            this.panel2.Controls.Add(this.rad_in);
            this.panel2.Controls.Add(this.txt_DTe);
            this.panel2.Controls.Add(this.label9);
            this.panel2.Controls.Add(this.txt_DTs);
            this.panel2.Controls.Add(this.label6);
            this.panel2.Controls.Add(this.txt_spec);
            this.panel2.Controls.Add(this.btn_qexport);
            this.panel2.Controls.Add(this.btn_delPosition);
            this.panel2.Controls.Add(this.btn_combi);
            this.panel2.Controls.Add(this.button1);
            this.panel2.Controls.Add(this.label2);
            this.panel2.Controls.Add(this.txt_position);
            this.panel2.Controls.Add(this.label1);
            this.panel2.Controls.Add(this.txt_item);
            this.panel2.Location = new System.Drawing.Point(65, 57);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(627, 70);
            this.panel2.TabIndex = 6;
            // 
            // rad_out
            // 
            this.rad_out.AutoSize = true;
            this.rad_out.Font = new System.Drawing.Font("新細明體", 8F);
            this.rad_out.Location = new System.Drawing.Point(438, 22);
            this.rad_out.Name = "rad_out";
            this.rad_out.Size = new System.Drawing.Size(34, 15);
            this.rad_out.TabIndex = 17;
            this.rad_out.Text = "出";
            this.rad_out.UseVisualStyleBackColor = true;
            // 
            // rad_in
            // 
            this.rad_in.AutoSize = true;
            this.rad_in.Checked = true;
            this.rad_in.Font = new System.Drawing.Font("新細明體", 8F);
            this.rad_in.Location = new System.Drawing.Point(398, 22);
            this.rad_in.Name = "rad_in";
            this.rad_in.Size = new System.Drawing.Size(34, 15);
            this.rad_in.TabIndex = 16;
            this.rad_in.TabStop = true;
            this.rad_in.Text = "入";
            this.rad_in.UseVisualStyleBackColor = true;
            // 
            // txt_DTe
            // 
            this.txt_DTe.Font = new System.Drawing.Font("微軟正黑體", 10F, System.Drawing.FontStyle.Bold);
            this.txt_DTe.Location = new System.Drawing.Point(387, 38);
            this.txt_DTe.Name = "txt_DTe";
            this.txt_DTe.Size = new System.Drawing.Size(85, 25);
            this.txt_DTe.TabIndex = 15;
            this.txt_DTe.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txt_DTs_KeyPress);
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Font = new System.Drawing.Font("微軟正黑體", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label9.Location = new System.Drawing.Point(236, 44);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(60, 17);
            this.label9.TabIndex = 14;
            this.label9.Text = "時間查詢";
            // 
            // txt_DTs
            // 
            this.txt_DTs.Font = new System.Drawing.Font("微軟正黑體", 10F, System.Drawing.FontStyle.Bold);
            this.txt_DTs.Location = new System.Drawing.Point(296, 38);
            this.txt_DTs.Name = "txt_DTs";
            this.txt_DTs.Size = new System.Drawing.Size(85, 25);
            this.txt_DTs.TabIndex = 13;
            this.txt_DTs.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txt_DTs_KeyPress);
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Font = new System.Drawing.Font("微軟正黑體", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label6.Location = new System.Drawing.Point(11, 44);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(60, 17);
            this.label6.TabIndex = 12;
            this.label6.Text = "規格查詢";
            // 
            // txt_spec
            // 
            this.txt_spec.Font = new System.Drawing.Font("微軟正黑體", 11.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.txt_spec.Location = new System.Drawing.Point(73, 38);
            this.txt_spec.Name = "txt_spec";
            this.txt_spec.Size = new System.Drawing.Size(157, 27);
            this.txt_spec.TabIndex = 11;
            this.txt_spec.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txt_spec_KeyPress);
            // 
            // btn_qexport
            // 
            this.btn_qexport.BackColor = System.Drawing.Color.OldLace;
            this.btn_qexport.Font = new System.Drawing.Font("微軟正黑體", 12F, System.Drawing.FontStyle.Bold);
            this.btn_qexport.Location = new System.Drawing.Point(569, 37);
            this.btn_qexport.Name = "btn_qexport";
            this.btn_qexport.Size = new System.Drawing.Size(53, 27);
            this.btn_qexport.TabIndex = 10;
            this.btn_qexport.Text = "匯出";
            this.btn_qexport.UseVisualStyleBackColor = false;
            this.btn_qexport.Click += new System.EventHandler(this.btn_qexport_Click);
            // 
            // btn_delPosition
            // 
            this.btn_delPosition.BackColor = System.Drawing.Color.Salmon;
            this.btn_delPosition.Font = new System.Drawing.Font("微軟正黑體", 12F, System.Drawing.FontStyle.Bold);
            this.btn_delPosition.Location = new System.Drawing.Point(569, 7);
            this.btn_delPosition.Name = "btn_delPosition";
            this.btn_delPosition.Size = new System.Drawing.Size(36, 27);
            this.btn_delPosition.TabIndex = 9;
            this.btn_delPosition.Text = "刪";
            this.btn_delPosition.UseVisualStyleBackColor = false;
            this.btn_delPosition.Click += new System.EventHandler(this.btndelPosition_Click);
            // 
            // btn_combi
            // 
            this.btn_combi.BackColor = System.Drawing.Color.GreenYellow;
            this.btn_combi.Font = new System.Drawing.Font("微軟正黑體", 11F, System.Drawing.FontStyle.Bold);
            this.btn_combi.Location = new System.Drawing.Point(473, 37);
            this.btn_combi.Name = "btn_combi";
            this.btn_combi.Size = new System.Drawing.Size(91, 27);
            this.btn_combi.TabIndex = 8;
            this.btn_combi.Text = "料號+儲位";
            this.btn_combi.UseVisualStyleBackColor = false;
            this.btn_combi.Click += new System.EventHandler(this.btn_combi_Click);
            // 
            // button1
            // 
            this.button1.BackColor = System.Drawing.Color.GreenYellow;
            this.button1.Font = new System.Drawing.Font("微軟正黑體", 11F, System.Drawing.FontStyle.Bold);
            this.button1.Location = new System.Drawing.Point(473, 5);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(78, 27);
            this.button1.TabIndex = 7;
            this.button1.Text = "查詢全部";
            this.button1.UseVisualStyleBackColor = false;
            this.button1.Click += new System.EventHandler(this.button1_Click_1);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("微軟正黑體", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label2.Location = new System.Drawing.Point(236, 12);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(60, 17);
            this.label2.TabIndex = 3;
            this.label2.Text = "儲位查詢";
            // 
            // txt_position
            // 
            this.txt_position.Font = new System.Drawing.Font("微軟正黑體", 11.25F, System.Drawing.FontStyle.Bold);
            this.txt_position.Location = new System.Drawing.Point(296, 6);
            this.txt_position.Name = "txt_position";
            this.txt_position.Size = new System.Drawing.Size(85, 27);
            this.txt_position.TabIndex = 2;
            this.txt_position.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txt_position_KeyPress);
            this.txt_position.MouseDoubleClick += new System.Windows.Forms.MouseEventHandler(this.txt_position_MouseDoubleClick);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("微軟正黑體", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label1.Location = new System.Drawing.Point(11, 12);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(60, 17);
            this.label1.TabIndex = 1;
            this.label1.Text = "料號查詢";
            // 
            // txt_item
            // 
            this.txt_item.Font = new System.Drawing.Font("微軟正黑體", 11.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.txt_item.Location = new System.Drawing.Point(73, 6);
            this.txt_item.Name = "txt_item";
            this.txt_item.Size = new System.Drawing.Size(157, 27);
            this.txt_item.TabIndex = 0;
            this.txt_item.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txt_item_KeyPress);
            // 
            // panel3
            // 
            this.panel3.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.panel3.Controls.Add(this.btn_Input);
            this.panel3.Controls.Add(this.btn_Out);
            this.panel3.Controls.Add(this.btn_BatIn);
            this.panel3.Controls.Add(this.btn_BatOut);
            this.panel3.Location = new System.Drawing.Point(12, 1);
            this.panel3.Name = "panel3";
            this.panel3.Size = new System.Drawing.Size(776, 52);
            this.panel3.TabIndex = 7;
            // 
            // panel4
            // 
            this.panel4.Controls.Add(this.btn_history);
            this.panel4.Controls.Add(this.btn_Maintain);
            this.panel4.Location = new System.Drawing.Point(695, 56);
            this.panel4.Name = "panel4";
            this.panel4.Size = new System.Drawing.Size(93, 70);
            this.panel4.TabIndex = 8;
            // 
            // btn_history
            // 
            this.btn_history.BackColor = System.Drawing.Color.LightCoral;
            this.btn_history.Font = new System.Drawing.Font("微軟正黑體", 12F, System.Drawing.FontStyle.Bold);
            this.btn_history.Location = new System.Drawing.Point(2, 34);
            this.btn_history.Name = "btn_history";
            this.btn_history.Size = new System.Drawing.Size(89, 34);
            this.btn_history.TabIndex = 1;
            this.btn_history.Text = "歷史查詢";
            this.btn_history.UseVisualStyleBackColor = false;
            this.btn_history.Click += new System.EventHandler(this.btn_history_Click);
            // 
            // btn_Maintain
            // 
            this.btn_Maintain.Font = new System.Drawing.Font("微軟正黑體", 12F, System.Drawing.FontStyle.Bold);
            this.btn_Maintain.Location = new System.Drawing.Point(2, 1);
            this.btn_Maintain.Name = "btn_Maintain";
            this.btn_Maintain.Size = new System.Drawing.Size(89, 34);
            this.btn_Maintain.TabIndex = 0;
            this.btn_Maintain.Text = "管理設定";
            this.btn_Maintain.UseVisualStyleBackColor = true;
            this.btn_Maintain.Click += new System.EventHandler(this.Maintain_Click);
            // 
            // fileSystemWatcher1
            // 
            this.fileSystemWatcher1.EnableRaisingEvents = true;
            this.fileSystemWatcher1.SynchronizingObject = this;
            // 
            // inputExcelBW
            // 
            this.inputExcelBW.DoWork += new System.ComponentModel.DoWorkEventHandler(this.inputExcelBW_DoWork);
            this.inputExcelBW.ProgressChanged += new System.ComponentModel.ProgressChangedEventHandler(this.inputExcelBW_ProgressChanged);
            this.inputExcelBW.RunWorkerCompleted += new System.ComponentModel.RunWorkerCompletedEventHandler(this.inputExcelBW_RunWorkerCompleted);
            // 
            // panel5
            // 
            this.panel5.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel5.Controls.Add(this.label8);
            this.panel5.Controls.Add(this.txt_wonoOut);
            this.panel5.Controls.Add(this.label7);
            this.panel5.Controls.Add(this.txt_specP5);
            this.panel5.Controls.Add(this.btn_reP2);
            this.panel5.Controls.Add(this.btn_itemSite);
            this.panel5.Controls.Add(this.btn_findAll);
            this.panel5.Controls.Add(this.label3);
            this.panel5.Controls.Add(this.txt_siteP5);
            this.panel5.Controls.Add(this.label4);
            this.panel5.Controls.Add(this.txt_itemP5);
            this.panel5.Location = new System.Drawing.Point(65, 57);
            this.panel5.Name = "panel5";
            this.panel5.Size = new System.Drawing.Size(627, 70);
            this.panel5.TabIndex = 9;
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Font = new System.Drawing.Font("微軟正黑體", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label8.Location = new System.Drawing.Point(257, 44);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(60, 17);
            this.label8.TabIndex = 13;
            this.label8.Text = "出庫工單";
            // 
            // txt_wonoOut
            // 
            this.txt_wonoOut.Font = new System.Drawing.Font("微軟正黑體", 11.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.txt_wonoOut.Location = new System.Drawing.Point(320, 38);
            this.txt_wonoOut.Name = "txt_wonoOut";
            this.txt_wonoOut.Size = new System.Drawing.Size(132, 27);
            this.txt_wonoOut.TabIndex = 12;
            this.txt_wonoOut.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txt_wonoOut_KeyPress);
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Font = new System.Drawing.Font("微軟正黑體", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label7.Location = new System.Drawing.Point(12, 44);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(60, 17);
            this.label7.TabIndex = 11;
            this.label7.Text = "規格查詢";
            // 
            // txt_specP5
            // 
            this.txt_specP5.Font = new System.Drawing.Font("微軟正黑體", 11.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.txt_specP5.Location = new System.Drawing.Point(75, 38);
            this.txt_specP5.Name = "txt_specP5";
            this.txt_specP5.Size = new System.Drawing.Size(175, 27);
            this.txt_specP5.TabIndex = 10;
            this.txt_specP5.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txt_spec_KeyPress);
            // 
            // btn_reP2
            // 
            this.btn_reP2.BackColor = System.Drawing.Color.Salmon;
            this.btn_reP2.Font = new System.Drawing.Font("微軟正黑體", 12F, System.Drawing.FontStyle.Bold);
            this.btn_reP2.Location = new System.Drawing.Point(569, 7);
            this.btn_reP2.Name = "btn_reP2";
            this.btn_reP2.Size = new System.Drawing.Size(36, 27);
            this.btn_reP2.TabIndex = 9;
            this.btn_reP2.Text = "返回";
            this.btn_reP2.UseVisualStyleBackColor = false;
            this.btn_reP2.Click += new System.EventHandler(this.btn_reP2_Click);
            // 
            // btn_itemSite
            // 
            this.btn_itemSite.BackColor = System.Drawing.Color.GreenYellow;
            this.btn_itemSite.Font = new System.Drawing.Font("微軟正黑體", 12F, System.Drawing.FontStyle.Bold);
            this.btn_itemSite.Location = new System.Drawing.Point(458, 37);
            this.btn_itemSite.Name = "btn_itemSite";
            this.btn_itemSite.Size = new System.Drawing.Size(102, 27);
            this.btn_itemSite.TabIndex = 8;
            this.btn_itemSite.Text = "料號+儲位";
            this.btn_itemSite.UseVisualStyleBackColor = false;
            this.btn_itemSite.Click += new System.EventHandler(this.btn_itemSite_Click);
            // 
            // btn_findAll
            // 
            this.btn_findAll.BackColor = System.Drawing.Color.GreenYellow;
            this.btn_findAll.Font = new System.Drawing.Font("微軟正黑體", 12F, System.Drawing.FontStyle.Bold);
            this.btn_findAll.Location = new System.Drawing.Point(458, 5);
            this.btn_findAll.Name = "btn_findAll";
            this.btn_findAll.Size = new System.Drawing.Size(89, 27);
            this.btn_findAll.TabIndex = 7;
            this.btn_findAll.Text = "查詢全部";
            this.btn_findAll.UseVisualStyleBackColor = false;
            this.btn_findAll.Click += new System.EventHandler(this.btn_findAll_Click);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("微軟正黑體", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label3.Location = new System.Drawing.Point(257, 9);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(60, 17);
            this.label3.TabIndex = 3;
            this.label3.Text = "儲位查詢";
            // 
            // txt_siteP5
            // 
            this.txt_siteP5.Font = new System.Drawing.Font("微軟正黑體", 11.25F, System.Drawing.FontStyle.Bold);
            this.txt_siteP5.Location = new System.Drawing.Point(319, 3);
            this.txt_siteP5.Name = "txt_siteP5";
            this.txt_siteP5.Size = new System.Drawing.Size(100, 27);
            this.txt_siteP5.TabIndex = 2;
            this.txt_siteP5.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txt_position_KeyPress);
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("微軟正黑體", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label4.Location = new System.Drawing.Point(12, 12);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(60, 17);
            this.label4.TabIndex = 1;
            this.label4.Text = "料號查詢";
            // 
            // txt_itemP5
            // 
            this.txt_itemP5.Font = new System.Drawing.Font("微軟正黑體", 11.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.txt_itemP5.Location = new System.Drawing.Point(75, 6);
            this.txt_itemP5.Name = "txt_itemP5";
            this.txt_itemP5.Size = new System.Drawing.Size(175, 27);
            this.txt_itemP5.TabIndex = 0;
            this.txt_itemP5.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txt_item_KeyPress);
            // 
            // outputExcelBW
            // 
            this.outputExcelBW.DoWork += new System.ComponentModel.DoWorkEventHandler(this.outputExcelBW_DoWork);
            this.outputExcelBW.RunWorkerCompleted += new System.ComponentModel.RunWorkerCompletedEventHandler(this.outputExcelBW_RunWorkerCompleted);
            // 
            // progressBar2
            // 
            this.progressBar2.Location = new System.Drawing.Point(192, 292);
            this.progressBar2.Name = "progressBar2";
            this.progressBar2.Size = new System.Drawing.Size(401, 18);
            this.progressBar2.TabIndex = 10;
            this.progressBar2.Visible = false;
            // 
            // panel6
            // 
            this.panel6.Controls.Add(this.txt_sumC);
            this.panel6.Controls.Add(this.label5);
            this.panel6.Location = new System.Drawing.Point(255, 486);
            this.panel6.Name = "panel6";
            this.panel6.Size = new System.Drawing.Size(162, 35);
            this.panel6.TabIndex = 11;
            // 
            // txt_sumC
            // 
            this.txt_sumC.Font = new System.Drawing.Font("微軟正黑體", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.txt_sumC.Location = new System.Drawing.Point(62, 3);
            this.txt_sumC.Name = "txt_sumC";
            this.txt_sumC.ReadOnly = true;
            this.txt_sumC.Size = new System.Drawing.Size(72, 25);
            this.txt_sumC.TabIndex = 1;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(3, 8);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(50, 12);
            this.label5.TabIndex = 0;
            this.label5.Text = "總數量 : ";
            // 
            // rad_E
            // 
            this.rad_E.AutoSize = true;
            this.rad_E.Checked = true;
            this.rad_E.Location = new System.Drawing.Point(3, 9);
            this.rad_E.Name = "rad_E";
            this.rad_E.Size = new System.Drawing.Size(35, 16);
            this.rad_E.TabIndex = 12;
            this.rad_E.TabStop = true;
            this.rad_E.Text = "昶";
            this.rad_E.UseVisualStyleBackColor = true;
            // 
            // rad_C
            // 
            this.rad_C.AutoSize = true;
            this.rad_C.Location = new System.Drawing.Point(3, 38);
            this.rad_C.Name = "rad_C";
            this.rad_C.Size = new System.Drawing.Size(35, 16);
            this.rad_C.TabIndex = 13;
            this.rad_C.Text = "客";
            this.rad_C.UseVisualStyleBackColor = true;
            // 
            // panel_rad
            // 
            this.panel_rad.Controls.Add(this.rad_C);
            this.panel_rad.Controls.Add(this.rad_E);
            this.panel_rad.Location = new System.Drawing.Point(21, 59);
            this.panel_rad.Name = "panel_rad";
            this.panel_rad.Size = new System.Drawing.Size(39, 65);
            this.panel_rad.TabIndex = 14;
            // 
            // btnNextPage
            // 
            this.btnNextPage.Location = new System.Drawing.Point(695, 459);
            this.btnNextPage.Name = "btnNextPage";
            this.btnNextPage.Size = new System.Drawing.Size(75, 23);
            this.btnNextPage.TabIndex = 15;
            this.btnNextPage.Text = "下一頁";
            this.btnNextPage.UseVisualStyleBackColor = true;
            this.btnNextPage.Visible = false;
            this.btnNextPage.Click += new System.EventHandler(this.btnNextPage_Click);
            // 
            // btnPreviousPage
            // 
            this.btnPreviousPage.Location = new System.Drawing.Point(31, 459);
            this.btnPreviousPage.Name = "btnPreviousPage";
            this.btnPreviousPage.Size = new System.Drawing.Size(75, 23);
            this.btnPreviousPage.TabIndex = 15;
            this.btnPreviousPage.Text = "上一頁";
            this.btnPreviousPage.UseVisualStyleBackColor = true;
            this.btnPreviousPage.Visible = false;
            this.btnPreviousPage.Click += new System.EventHandler(this.btnPreviousPage_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(96F, 96F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi;
            this.BackColor = System.Drawing.Color.White;
            this.ClientSize = new System.Drawing.Size(800, 515);
            this.Controls.Add(this.btnPreviousPage);
            this.Controls.Add(this.btnNextPage);
            this.Controls.Add(this.panel_rad);
            this.Controls.Add(this.panel6);
            this.Controls.Add(this.panel2);
            this.Controls.Add(this.progressBar2);
            this.Controls.Add(this.panel5);
            this.Controls.Add(this.panel4);
            this.Controls.Add(this.panel3);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.dataGridView1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "Form1";
            this.Text = "儲位管理系統";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.Shown += new System.EventHandler(this.Form1_Shown);
            this.Resize += new System.EventHandler(this.Form1_Resize);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.panel2.ResumeLayout(false);
            this.panel2.PerformLayout();
            this.panel3.ResumeLayout(false);
            this.panel4.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.fileSystemWatcher1)).EndInit();
            this.panel5.ResumeLayout(false);
            this.panel5.PerformLayout();
            this.panel6.ResumeLayout(false);
            this.panel6.PerformLayout();
            this.panel_rad.ResumeLayout(false);
            this.panel_rad.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.Button btn_Input;
        private System.Windows.Forms.Button btn_Out;
        private System.Windows.Forms.Button btn_BatIn;
        private System.Windows.Forms.Button btn_BatOut;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.TextBox txt_path;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox txt_position;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox txt_item;
        private System.Windows.Forms.Button commitButton;
        private System.Windows.Forms.Panel panel3;
        private System.Windows.Forms.Panel panel4;
        private System.Windows.Forms.Button btn_Maintain;
        private System.IO.FileSystemWatcher fileSystemWatcher1;
        private System.Windows.Forms.Button btn_combi;
        private System.Windows.Forms.Button btn_delPosition;
        private System.Windows.Forms.Button btn_return;
        private System.Windows.Forms.ListBox list_result;
        private System.ComponentModel.BackgroundWorker inputExcelBW;
        private System.Windows.Forms.ProgressBar progressBar1;
        private System.Windows.Forms.Button selectButton;
        private System.Windows.Forms.Button btn_history;
        private System.Windows.Forms.Panel panel5;
        private System.Windows.Forms.Button btn_reP2;
        private System.Windows.Forms.Button btn_itemSite;
        private System.Windows.Forms.Button btn_findAll;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox txt_siteP5;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox txt_itemP5;
        private System.Windows.Forms.Button btn_qexport;
        private System.ComponentModel.BackgroundWorker outputExcelBW;
        private System.Windows.Forms.ProgressBar progressBar2;
        private System.Windows.Forms.Panel panel6;
        private System.Windows.Forms.TextBox txt_sumC;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.RadioButton rad_C;
        private System.Windows.Forms.RadioButton rad_E;
        private System.Windows.Forms.Panel panel_rad;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.TextBox txt_spec;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.TextBox txt_specP5;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.TextBox txt_wonoOut;
        private System.Windows.Forms.TextBox txt_result;
        private System.Windows.Forms.TextBox txt_DTs;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.TextBox txt_DTe;
        private System.Windows.Forms.RadioButton rad_out;
        private System.Windows.Forms.RadioButton rad_in;
        private System.Windows.Forms.Button btnPreviousPage;
        private System.Windows.Forms.Button btnNextPage;
    }
}

