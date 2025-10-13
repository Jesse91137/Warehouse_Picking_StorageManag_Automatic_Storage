
namespace Automatic_Storage
{
    partial class BatOut
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.panel1 = new System.Windows.Forms.Panel();
            this.txt_engsr = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.commitButton2 = new System.Windows.Forms.Button();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.txt_wono = new System.Windows.Forms.TextBox();
            this.btn_export = new System.Windows.Forms.Button();
            this.label3 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.selectButton = new System.Windows.Forms.Button();
            this.commitButton = new System.Windows.Forms.Button();
            this.txt_path = new System.Windows.Forms.TextBox();
            this.btn_Out = new System.Windows.Forms.Button();
            this.txt_item2 = new System.Windows.Forms.TextBox();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.loadExcelBW = new System.ComponentModel.BackgroundWorker();
            this.expExcelBW = new System.ComponentModel.BackgroundWorker();
            this.panel2 = new System.Windows.Forms.Panel();
            this.txt_errLog = new System.Windows.Forms.TextBox();
            this.label7 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.txt_source = new System.Windows.Forms.TextBox();
            this.txt_newPosition = new System.Windows.Forms.TextBox();
            this.btn_change = new System.Windows.Forms.Button();
            this.panel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.panel2.SuspendLayout();
            this.SuspendLayout();
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.txt_engsr);
            this.panel1.Controls.Add(this.label2);
            this.panel1.Controls.Add(this.commitButton2);
            this.panel1.Controls.Add(this.progressBar1);
            this.panel1.Controls.Add(this.txt_wono);
            this.panel1.Controls.Add(this.btn_export);
            this.panel1.Controls.Add(this.label3);
            this.panel1.Controls.Add(this.label1);
            this.panel1.Controls.Add(this.selectButton);
            this.panel1.Controls.Add(this.commitButton);
            this.panel1.Controls.Add(this.txt_path);
            this.panel1.Location = new System.Drawing.Point(12, 3);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(627, 116);
            this.panel1.TabIndex = 6;
            // 
            // txt_engsr
            // 
            this.txt_engsr.Location = new System.Drawing.Point(271, 9);
            this.txt_engsr.Name = "txt_engsr";
            this.txt_engsr.Size = new System.Drawing.Size(132, 22);
            this.txt_engsr.TabIndex = 18;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(211, 12);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(53, 12);
            this.label2.TabIndex = 19;
            this.label2.Text = "機種名稱";
            // 
            // commitButton2
            // 
            this.commitButton2.Font = new System.Drawing.Font("微軟正黑體", 12F, System.Drawing.FontStyle.Bold);
            this.commitButton2.Location = new System.Drawing.Point(508, 76);
            this.commitButton2.Name = "commitButton2";
            this.commitButton2.Size = new System.Drawing.Size(89, 30);
            this.commitButton2.TabIndex = 17;
            this.commitButton2.Text = "A4單";
            this.commitButton2.UseVisualStyleBackColor = true;
            this.commitButton2.Click += new System.EventHandler(this.commitButton2_Click);
            // 
            // progressBar1
            // 
            this.progressBar1.Location = new System.Drawing.Point(5, 85);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(363, 10);
            this.progressBar1.TabIndex = 13;
            // 
            // txt_wono
            // 
            this.txt_wono.Location = new System.Drawing.Point(63, 9);
            this.txt_wono.Name = "txt_wono";
            this.txt_wono.Size = new System.Drawing.Size(132, 22);
            this.txt_wono.TabIndex = 2;
            this.txt_wono.Leave += new System.EventHandler(this.txt_wono_Leave);
            // 
            // btn_export
            // 
            this.btn_export.Font = new System.Drawing.Font("微軟正黑體", 12F, System.Drawing.FontStyle.Bold);
            this.btn_export.Location = new System.Drawing.Point(398, 76);
            this.btn_export.Name = "btn_export";
            this.btn_export.Size = new System.Drawing.Size(89, 30);
            this.btn_export.TabIndex = 12;
            this.btn_export.Text = "資料匯出";
            this.btn_export.UseVisualStyleBackColor = true;
            this.btn_export.Click += new System.EventHandler(this.btn_export_Click);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(3, 12);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(53, 12);
            this.label3.TabIndex = 10;
            this.label3.Text = "工單號碼";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(3, 48);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(53, 12);
            this.label1.TabIndex = 10;
            this.label1.Text = "選擇檔案";
            // 
            // selectButton
            // 
            this.selectButton.Font = new System.Drawing.Font("微軟正黑體", 12F, System.Drawing.FontStyle.Bold);
            this.selectButton.Location = new System.Drawing.Point(398, 40);
            this.selectButton.Name = "selectButton";
            this.selectButton.Size = new System.Drawing.Size(89, 30);
            this.selectButton.TabIndex = 8;
            this.selectButton.Text = "檔案選擇";
            this.selectButton.UseVisualStyleBackColor = true;
            this.selectButton.Click += new System.EventHandler(this.selectButton_Click);
            // 
            // commitButton
            // 
            this.commitButton.Font = new System.Drawing.Font("微軟正黑體", 12F, System.Drawing.FontStyle.Bold);
            this.commitButton.Location = new System.Drawing.Point(508, 40);
            this.commitButton.Name = "commitButton";
            this.commitButton.Size = new System.Drawing.Size(89, 30);
            this.commitButton.TabIndex = 8;
            this.commitButton.Text = "SAP單";
            this.commitButton.UseVisualStyleBackColor = true;
            this.commitButton.Click += new System.EventHandler(this.commitButton_Click);
            // 
            // txt_path
            // 
            this.txt_path.Location = new System.Drawing.Point(63, 45);
            this.txt_path.Name = "txt_path";
            this.txt_path.Size = new System.Drawing.Size(304, 22);
            this.txt_path.TabIndex = 3;
            // 
            // btn_Out
            // 
            this.btn_Out.BackColor = System.Drawing.Color.MistyRose;
            this.btn_Out.Enabled = false;
            this.btn_Out.Font = new System.Drawing.Font("微軟正黑體", 12F, System.Drawing.FontStyle.Bold);
            this.btn_Out.Location = new System.Drawing.Point(298, 451);
            this.btn_Out.Name = "btn_Out";
            this.btn_Out.Size = new System.Drawing.Size(89, 30);
            this.btn_Out.TabIndex = 14;
            this.btn_Out.Text = "出庫登記";
            this.btn_Out.UseVisualStyleBackColor = false;
            this.btn_Out.Visible = false;
            this.btn_Out.Click += new System.EventHandler(this.btn_Out_Click);
            // 
            // txt_item2
            // 
            this.txt_item2.Enabled = false;
            this.txt_item2.Location = new System.Drawing.Point(393, 451);
            this.txt_item2.Name = "txt_item2";
            this.txt_item2.Size = new System.Drawing.Size(244, 22);
            this.txt_item2.TabIndex = 4;
            this.txt_item2.Visible = false;
            // 
            // dataGridView1
            // 
            this.dataGridView1.AllowUserToAddRows = false;
            this.dataGridView1.AllowUserToDeleteRows = false;
            this.dataGridView1.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this.dataGridView1.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.DisplayedCells;
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Location = new System.Drawing.Point(12, 123);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.ReadOnly = true;
            this.dataGridView1.RowTemplate.Height = 24;
            this.dataGridView1.Size = new System.Drawing.Size(627, 322);
            this.dataGridView1.TabIndex = 7;
            this.dataGridView1.DataBindingComplete += new System.Windows.Forms.DataGridViewBindingCompleteEventHandler(this.dataGridView1_DataBindingComplete);
            this.dataGridView1.RowStateChanged += new System.Windows.Forms.DataGridViewRowStateChangedEventHandler(this.dataGridView1_RowStateChanged);
            this.dataGridView1.Scroll += new System.Windows.Forms.ScrollEventHandler(this.dataGridView1_Scroll);
            this.dataGridView1.DoubleClick += new System.EventHandler(this.dataGridView1_DoubleClick);
            // 
            // loadExcelBW
            // 
            this.loadExcelBW.DoWork += new System.ComponentModel.DoWorkEventHandler(this.loadExcelBW_DoWork);
            this.loadExcelBW.ProgressChanged += new System.ComponentModel.ProgressChangedEventHandler(this.loadExcelBW_ProgressChanged);
            // 
            // expExcelBW
            // 
            this.expExcelBW.DoWork += new System.ComponentModel.DoWorkEventHandler(this.expExcelBW_DoWork);
            this.expExcelBW.ProgressChanged += new System.ComponentModel.ProgressChangedEventHandler(this.expExcelBW_ProgressChanged);
            this.expExcelBW.RunWorkerCompleted += new System.ComponentModel.RunWorkerCompletedEventHandler(this.expExcelBW_RunWorkerCompleted);
            // 
            // panel2
            // 
            this.panel2.Controls.Add(this.txt_errLog);
            this.panel2.Controls.Add(this.label7);
            this.panel2.Controls.Add(this.label5);
            this.panel2.Controls.Add(this.label4);
            this.panel2.Controls.Add(this.txt_source);
            this.panel2.Controls.Add(this.txt_newPosition);
            this.panel2.Controls.Add(this.btn_change);
            this.panel2.Location = new System.Drawing.Point(12, 3);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(627, 442);
            this.panel2.TabIndex = 15;
            // 
            // txt_errLog
            // 
            this.txt_errLog.Location = new System.Drawing.Point(76, 238);
            this.txt_errLog.Multiline = true;
            this.txt_errLog.Name = "txt_errLog";
            this.txt_errLog.Size = new System.Drawing.Size(354, 145);
            this.txt_errLog.TabIndex = 23;
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Font = new System.Drawing.Font("微軟正黑體", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label7.ForeColor = System.Drawing.Color.SteelBlue;
            this.label7.Location = new System.Drawing.Point(72, 180);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(105, 24);
            this.label7.TabIndex = 22;
            this.label7.Text = "輸入新儲位";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Font = new System.Drawing.Font("微軟正黑體", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label5.ForeColor = System.Drawing.Color.SteelBlue;
            this.label5.Location = new System.Drawing.Point(72, 120);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(124, 24);
            this.label5.TabIndex = 20;
            this.label5.Text = "輸入原本儲位";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("微軟正黑體", 20.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label4.ForeColor = System.Drawing.Color.Coral;
            this.label4.Location = new System.Drawing.Point(118, 45);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(312, 35);
            this.label4.TabIndex = 19;
            this.label4.Text = "依據儲位查詢後整批異動";
            // 
            // txt_source
            // 
            this.txt_source.Location = new System.Drawing.Point(213, 122);
            this.txt_source.Name = "txt_source";
            this.txt_source.Size = new System.Drawing.Size(217, 22);
            this.txt_source.TabIndex = 16;
            // 
            // txt_newPosition
            // 
            this.txt_newPosition.Location = new System.Drawing.Point(213, 180);
            this.txt_newPosition.Name = "txt_newPosition";
            this.txt_newPosition.Size = new System.Drawing.Size(217, 22);
            this.txt_newPosition.TabIndex = 17;
            // 
            // btn_change
            // 
            this.btn_change.Location = new System.Drawing.Point(481, 178);
            this.btn_change.Name = "btn_change";
            this.btn_change.Size = new System.Drawing.Size(73, 26);
            this.btn_change.TabIndex = 18;
            this.btn_change.Text = "儲位變更";
            this.btn_change.UseVisualStyleBackColor = true;
            this.btn_change.Click += new System.EventHandler(this.btn_change_Click);
            // 
            // BatOut
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(649, 451);
            this.Controls.Add(this.panel2);
            this.Controls.Add(this.dataGridView1);
            this.Controls.Add(this.btn_Out);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.txt_item2);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Name = "BatOut";
            this.Text = "批次出庫";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.BatOut_FormClosing);
            this.Load += new System.EventHandler(this.BatOut_Load);
            this.Shown += new System.EventHandler(this.BatOut_Shown);
            this.Resize += new System.EventHandler(this.BatOut_Resize);
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.panel2.ResumeLayout(false);
            this.panel2.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Button commitButton;
        private System.Windows.Forms.TextBox txt_path;
        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.TextBox txt_item2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button btn_export;
        private System.Windows.Forms.TextBox txt_wono;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.ProgressBar progressBar1;
        private System.ComponentModel.BackgroundWorker loadExcelBW;
        private System.ComponentModel.BackgroundWorker expExcelBW;
        private System.Windows.Forms.Button selectButton;
        private System.Windows.Forms.Button btn_Out;
        private System.Windows.Forms.Button commitButton2;
        private System.Windows.Forms.TextBox txt_engsr;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.TextBox txt_source;
        private System.Windows.Forms.TextBox txt_newPosition;
        private System.Windows.Forms.Button btn_change;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox txt_errLog;
    }
}