
namespace Automatic_Storage
{
    partial class OutPut
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(OutPut));
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label11 = new System.Windows.Forms.Label();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.textBox2 = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.textBox3 = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.loadExcelBW = new System.ComponentModel.BackgroundWorker();
            this.expExcelBW = new System.ComponentModel.BackgroundWorker();
            this.background = new System.ComponentModel.BackgroundWorker();
            this.label7 = new System.Windows.Forms.Label();
            this.textBox4 = new System.Windows.Forms.TextBox();
            this.cbx_package = new System.Windows.Forms.ComboBox();
            this.label9 = new System.Windows.Forms.Label();
            this.label10 = new System.Windows.Forms.Label();
            this.bat_confirm = new System.Windows.Forms.Button();
            this.label8 = new System.Windows.Forms.Label();
            this.txt_itemC = new System.Windows.Forms.TextBox();
            this.label6 = new System.Windows.Forms.Label();
            this.txt_Mark = new System.Windows.Forms.TextBox();
            this.visible_Panel = new System.Windows.Forms.Panel();
            this.label12 = new System.Windows.Forms.Label();
            this.txt_reelid = new System.Windows.Forms.TextBox();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.SuspendLayout();
            // 
            // textBox1
            // 
            this.textBox1.Font = new System.Drawing.Font("微軟正黑體", 12F, System.Drawing.FontStyle.Bold);
            this.textBox1.Location = new System.Drawing.Point(133, 12);
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(284, 29);
            this.textBox1.TabIndex = 50;
            this.textBox1.TextChanged += new System.EventHandler(this.textBox1_TextChanged);
            this.textBox1.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.textBox1_KeyPress);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("微軟正黑體", 12F, System.Drawing.FontStyle.Bold);
            this.label1.Location = new System.Drawing.Point(34, 16);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(74, 21);
            this.label1.TabIndex = 1;
            this.label1.Text = "昶亨料號";
            // 
            // label11
            // 
            this.label11.AutoSize = true;
            this.label11.Location = new System.Drawing.Point(2, 88);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(419, 12);
            this.label11.TabIndex = 15;
            this.label11.Text = "_____________________________________________________________________";
            // 
            // dataGridView1
            // 
            this.dataGridView1.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.DisplayedCells;
            this.dataGridView1.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.DisplayedCells;
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Location = new System.Drawing.Point(14, 110);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.RowTemplate.Height = 24;
            this.dataGridView1.Size = new System.Drawing.Size(403, 113);
            this.dataGridView1.TabIndex = 99;
            this.dataGridView1.Visible = false;
            this.dataGridView1.DataBindingComplete += new System.Windows.Forms.DataGridViewBindingCompleteEventHandler(this.dataGridView1_DataBindingComplete);
            this.dataGridView1.DoubleClick += new System.EventHandler(this.dataGridView1_DoubleClick);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(2, 225);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(419, 12);
            this.label2.TabIndex = 19;
            this.label2.Text = "_____________________________________________________________________";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("微軟正黑體", 12F, System.Drawing.FontStyle.Bold);
            this.label3.Location = new System.Drawing.Point(4, 434);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(74, 21);
            this.label3.TabIndex = 18;
            this.label3.Text = "料號確認";
            // 
            // textBox2
            // 
            this.textBox2.Font = new System.Drawing.Font("微軟正黑體", 12F, System.Drawing.FontStyle.Bold);
            this.textBox2.Location = new System.Drawing.Point(134, 431);
            this.textBox2.Name = "textBox2";
            this.textBox2.Size = new System.Drawing.Size(216, 29);
            this.textBox2.TabIndex = 51;
            this.textBox2.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.textBox2_KeyPress);
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("微軟正黑體", 12F, System.Drawing.FontStyle.Bold);
            this.label4.Location = new System.Drawing.Point(34, 396);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(74, 21);
            this.label4.TabIndex = 21;
            this.label4.Text = "儲位確認";
            // 
            // textBox3
            // 
            this.textBox3.Font = new System.Drawing.Font("微軟正黑體", 12F, System.Drawing.FontStyle.Bold);
            this.textBox3.Location = new System.Drawing.Point(133, 392);
            this.textBox3.Name = "textBox3";
            this.textBox3.Size = new System.Drawing.Size(186, 29);
            this.textBox3.TabIndex = 52;
            this.textBox3.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.textBox3_KeyPress);
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Font = new System.Drawing.Font("微軟正黑體", 20.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label5.Location = new System.Drawing.Point(64, 169);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(0, 35);
            this.label5.TabIndex = 100;
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Font = new System.Drawing.Font("微軟正黑體", 12F, System.Drawing.FontStyle.Bold);
            this.label7.Location = new System.Drawing.Point(4, 250);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(107, 21);
            this.label7.TabIndex = 102;
            this.label7.Text = "出庫數量/Pcs";
            // 
            // textBox4
            // 
            this.textBox4.Font = new System.Drawing.Font("微軟正黑體", 12F, System.Drawing.FontStyle.Bold);
            this.textBox4.Location = new System.Drawing.Point(133, 247);
            this.textBox4.Name = "textBox4";
            this.textBox4.Size = new System.Drawing.Size(151, 29);
            this.textBox4.TabIndex = 103;
            this.textBox4.Text = "1";
            this.textBox4.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.textBox4_KeyPress);
            this.textBox4.Leave += new System.EventHandler(this.textBox4_Leave);
            // 
            // cbx_package
            // 
            this.cbx_package.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbx_package.Font = new System.Drawing.Font("微軟正黑體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.cbx_package.FormattingEnabled = true;
            this.cbx_package.ItemHeight = 20;
            this.cbx_package.Location = new System.Drawing.Point(133, 284);
            this.cbx_package.Name = "cbx_package";
            this.cbx_package.Size = new System.Drawing.Size(216, 28);
            this.cbx_package.TabIndex = 106;
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Font = new System.Drawing.Font("微軟正黑體", 12F, System.Drawing.FontStyle.Bold);
            this.label9.Location = new System.Drawing.Point(34, 287);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(74, 21);
            this.label9.TabIndex = 107;
            this.label9.Text = "包裝種類";
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Font = new System.Drawing.Font("微軟正黑體", 12F, System.Drawing.FontStyle.Bold);
            this.label10.ForeColor = System.Drawing.Color.Red;
            this.label10.Location = new System.Drawing.Point(34, 473);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(0, 21);
            this.label10.TabIndex = 108;
            // 
            // bat_confirm
            // 
            this.bat_confirm.Font = new System.Drawing.Font("微軟正黑體", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.bat_confirm.Location = new System.Drawing.Point(354, 430);
            this.bat_confirm.Name = "bat_confirm";
            this.bat_confirm.Size = new System.Drawing.Size(61, 42);
            this.bat_confirm.TabIndex = 109;
            this.bat_confirm.Text = "批次出庫確認";
            this.bat_confirm.UseVisualStyleBackColor = true;
            this.bat_confirm.Visible = false;
            this.bat_confirm.Click += new System.EventHandler(this.bat_confirm_Click);
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Font = new System.Drawing.Font("微軟正黑體", 12F, System.Drawing.FontStyle.Bold);
            this.label8.Location = new System.Drawing.Point(34, 58);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(74, 21);
            this.label8.TabIndex = 110;
            this.label8.Text = "客戶料號";
            // 
            // txt_itemC
            // 
            this.txt_itemC.Font = new System.Drawing.Font("微軟正黑體", 12F, System.Drawing.FontStyle.Bold);
            this.txt_itemC.Location = new System.Drawing.Point(133, 54);
            this.txt_itemC.Name = "txt_itemC";
            this.txt_itemC.Size = new System.Drawing.Size(284, 29);
            this.txt_itemC.TabIndex = 51;
            this.txt_itemC.TextChanged += new System.EventHandler(this.textBox1_TextChanged);
            this.txt_itemC.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.textBox1_KeyPress);
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Font = new System.Drawing.Font("微軟正黑體", 12F, System.Drawing.FontStyle.Bold);
            this.label6.Location = new System.Drawing.Point(66, 321);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(42, 21);
            this.label6.TabIndex = 112;
            this.label6.Text = "備註";
            // 
            // txt_Mark
            // 
            this.txt_Mark.Font = new System.Drawing.Font("微軟正黑體", 12F, System.Drawing.FontStyle.Bold);
            this.txt_Mark.Location = new System.Drawing.Point(133, 318);
            this.txt_Mark.Name = "txt_Mark";
            this.txt_Mark.Size = new System.Drawing.Size(216, 29);
            this.txt_Mark.TabIndex = 111;
            // 
            // visible_Panel
            // 
            this.visible_Panel.Location = new System.Drawing.Point(4, 431);
            this.visible_Panel.Name = "visible_Panel";
            this.visible_Panel.Size = new System.Drawing.Size(346, 39);
            this.visible_Panel.TabIndex = 113;
            // 
            // label12
            // 
            this.label12.AutoSize = true;
            this.label12.Font = new System.Drawing.Font("微軟正黑體", 12F, System.Drawing.FontStyle.Bold);
            this.label12.Location = new System.Drawing.Point(34, 358);
            this.label12.Name = "label12";
            this.label12.Size = new System.Drawing.Size(67, 21);
            this.label12.TabIndex = 114;
            this.label12.Text = "Reel_ID";
            // 
            // txt_reelid
            // 
            this.txt_reelid.Font = new System.Drawing.Font("微軟正黑體", 12F, System.Drawing.FontStyle.Bold);
            this.txt_reelid.Location = new System.Drawing.Point(133, 354);
            this.txt_reelid.Name = "txt_reelid";
            this.txt_reelid.Size = new System.Drawing.Size(186, 29);
            this.txt_reelid.TabIndex = 115;
            // 
            // OutPut
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(429, 514);
            this.Controls.Add(this.label12);
            this.Controls.Add(this.txt_reelid);
            this.Controls.Add(this.visible_Panel);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.txt_Mark);
            this.Controls.Add(this.label8);
            this.Controls.Add(this.txt_itemC);
            this.Controls.Add(this.bat_confirm);
            this.Controls.Add(this.label10);
            this.Controls.Add(this.cbx_package);
            this.Controls.Add(this.label9);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.textBox4);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.textBox3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.textBox2);
            this.Controls.Add(this.dataGridView1);
            this.Controls.Add(this.label11);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.textBox1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximumSize = new System.Drawing.Size(640, 800);
            this.Name = "OutPut";
            this.Text = "出庫登記";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.OutPut_FormClosing);
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.OutPut_FormClosed);
            this.Load += new System.EventHandler(this.OutPut_Load);
            this.Shown += new System.EventHandler(this.OutPut_Shown);
            this.Resize += new System.EventHandler(this.OutPut_Resize);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.ComponentModel.BackgroundWorker background;
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label11;
        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox textBox2;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox textBox3;
        private System.Windows.Forms.Label label5;
        private System.ComponentModel.BackgroundWorker loadExcelBW;
        private System.ComponentModel.BackgroundWorker expExcelBW;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.TextBox textBox4;
        private System.Windows.Forms.ComboBox cbx_package;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.Button bat_confirm;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.TextBox txt_itemC;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.TextBox txt_Mark;
        private System.Windows.Forms.Panel visible_Panel;
        private System.Windows.Forms.Label label12;
        private System.Windows.Forms.TextBox txt_reelid;
    }
}