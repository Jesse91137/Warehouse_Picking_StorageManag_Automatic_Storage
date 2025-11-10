namespace Automatic_Storage
{
    partial class Form備料單匯入
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
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle2 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle3 = new System.Windows.Forms.DataGridViewCellStyle();
            this.lbl料號 = new System.Windows.Forms.Label();
            this.lbl數量 = new System.Windows.Forms.Label();
            this.dgv備料單 = new System.Windows.Forms.DataGridView();
            this.btn備料單匯入檔案 = new System.Windows.Forms.Button();
            this.txt備料單料號 = new System.Windows.Forms.TextBox();
            this.txt備料單數量 = new System.Windows.Forms.TextBox();
            this.btn備料單匯出 = new System.Windows.Forms.Button();
            this.btn備料單Unlock = new System.Windows.Forms.Button();
            this.btn備料單返回 = new System.Windows.Forms.Button();
            this.pnlTop = new System.Windows.Forms.Panel();
            this.btn備料單存檔 = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.dgv備料單)).BeginInit();
            this.pnlTop.SuspendLayout();
            this.SuspendLayout();
            // 
            // lbl料號
            // 
            this.lbl料號.AutoSize = true;
            this.lbl料號.Font = new System.Drawing.Font("微軟正黑體", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.lbl料號.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(50)))), ((int)(((byte)(50)))), ((int)(((byte)(50)))));
            this.lbl料號.Location = new System.Drawing.Point(22, 22);
            this.lbl料號.Name = "lbl料號";
            this.lbl料號.Size = new System.Drawing.Size(72, 25);
            this.lbl料號.TabIndex = 0;
            this.lbl料號.Text = "料號：";
            // 
            // lbl數量
            // 
            this.lbl數量.AutoSize = true;
            this.lbl數量.Font = new System.Drawing.Font("微軟正黑體", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.lbl數量.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(50)))), ((int)(((byte)(50)))), ((int)(((byte)(50)))));
            this.lbl數量.Location = new System.Drawing.Point(353, 22);
            this.lbl數量.Name = "lbl數量";
            this.lbl數量.Size = new System.Drawing.Size(72, 25);
            this.lbl數量.TabIndex = 0;
            this.lbl數量.Text = "數量：";
            // 
            // dgv備料單
            // 
            this.dgv備料單.AllowUserToAddRows = false;
            this.dgv備料單.AllowUserToDeleteRows = false;
            this.dgv備料單.AllowUserToResizeRows = false;
            dataGridViewCellStyle1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(250)))), ((int)(((byte)(250)))), ((int)(((byte)(250)))));
            dataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.dgv備料單.AlternatingRowsDefaultCellStyle = dataGridViewCellStyle1;
            this.dgv備料單.BackgroundColor = System.Drawing.Color.WhiteSmoke;
            dataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle2.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(240)))), ((int)(((byte)(240)))), ((int)(((byte)(240)))));
            dataGridViewCellStyle2.Font = new System.Drawing.Font("微軟正黑體", 12F, System.Drawing.FontStyle.Bold);
            dataGridViewCellStyle2.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(40)))), ((int)(((byte)(40)))), ((int)(((byte)(40)))));
            dataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.dgv備料單.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle2;
            this.dgv備料單.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle3.BackColor = System.Drawing.SystemColors.Window;
            dataGridViewCellStyle3.Font = new System.Drawing.Font("微軟正黑體", 12F, System.Drawing.FontStyle.Bold);
            dataGridViewCellStyle3.ForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle3.SelectionBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(180)))), ((int)(((byte)(215)))), ((int)(((byte)(255)))));
            dataGridViewCellStyle3.SelectionForeColor = System.Drawing.Color.Black;
            dataGridViewCellStyle3.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.dgv備料單.DefaultCellStyle = dataGridViewCellStyle3;
            this.dgv備料單.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dgv備料單.EnableHeadersVisualStyles = false;
            this.dgv備料單.GridColor = System.Drawing.Color.LightGray;
            this.dgv備料單.Location = new System.Drawing.Point(0, 70);
            this.dgv備料單.Name = "dgv備料單";
            this.dgv備料單.ReadOnly = true;
            this.dgv備料單.RowHeadersVisible = false;
            this.dgv備料單.RowHeadersWidth = 51;
            this.dgv備料單.RowTemplate.Height = 24;
            this.dgv備料單.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dgv備料單.Size = new System.Drawing.Size(1307, 652);
            this.dgv備料單.TabIndex = 0;
            this.dgv備料單.DataBindingComplete += new System.Windows.Forms.DataGridViewBindingCompleteEventHandler(this.Dgv備料單_DataBindingComplete);
            // 
            // btn備料單匯入檔案
            // 
            this.btn備料單匯入檔案.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(45)))), ((int)(((byte)(137)))), ((int)(((byte)(239)))));
            this.btn備料單匯入檔案.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btn備料單匯入檔案.FlatAppearance.BorderSize = 0;
            this.btn備料單匯入檔案.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btn備料單匯入檔案.Font = new System.Drawing.Font("微軟正黑體", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.btn備料單匯入檔案.ForeColor = System.Drawing.Color.White;
            this.btn備料單匯入檔案.Location = new System.Drawing.Point(583, 13);
            this.btn備料單匯入檔案.Name = "btn備料單匯入檔案";
            this.btn備料單匯入檔案.Size = new System.Drawing.Size(111, 38);
            this.btn備料單匯入檔案.TabIndex = 1;
            this.btn備料單匯入檔案.Text = "匯入檔案";
            this.btn備料單匯入檔案.UseVisualStyleBackColor = false;
            this.btn備料單匯入檔案.Click += new System.EventHandler(this.btn備料單匯入檔案_Click);
            // 
            // txt備料單料號
            // 
            this.txt備料單料號.Font = new System.Drawing.Font("微軟正黑體", 11F);
            this.txt備料單料號.Location = new System.Drawing.Point(100, 19);
            this.txt備料單料號.Name = "txt備料單料號";
            this.txt備料單料號.Size = new System.Drawing.Size(222, 32);
            this.txt備料單料號.TabIndex = 2;
            this.txt備料單料號.KeyDown += new System.Windows.Forms.KeyEventHandler(this.Txt備料單料號_KeyDown);
            // 
            // txt備料單數量
            // 
            this.txt備料單數量.Font = new System.Drawing.Font("微軟正黑體", 11F);
            this.txt備料單數量.Location = new System.Drawing.Point(431, 19);
            this.txt備料單數量.Name = "txt備料單數量";
            this.txt備料單數量.Size = new System.Drawing.Size(77, 32);
            this.txt備料單數量.TabIndex = 3;
            this.txt備料單數量.KeyDown += new System.Windows.Forms.KeyEventHandler(this.Txt備料單數量_KeyDown);
            this.txt備料單數量.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.Txt備料單數量_KeyPress);
            // 
            // btn備料單匯出
            // 
            this.btn備料單匯出.BackColor = System.Drawing.Color.Salmon;
            this.btn備料單匯出.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btn備料單匯出.FlatAppearance.BorderSize = 0;
            this.btn備料單匯出.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btn備料單匯出.Font = new System.Drawing.Font("微軟正黑體", 12F, System.Drawing.FontStyle.Bold);
            this.btn備料單匯出.ForeColor = System.Drawing.Color.Black;
            this.btn備料單匯出.Location = new System.Drawing.Point(844, 13);
            this.btn備料單匯出.Name = "btn備料單匯出";
            this.btn備料單匯出.Size = new System.Drawing.Size(124, 38);
            this.btn備料單匯出.TabIndex = 5;
            this.btn備料單匯出.Text = "匯出檔案";
            this.btn備料單匯出.UseVisualStyleBackColor = false;
            this.btn備料單匯出.Click += new System.EventHandler(this.btn備料單匯出_Click);
            // 
            // btn備料單Unlock
            // 
            this.btn備料單Unlock.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(45)))), ((int)(((byte)(137)))), ((int)(((byte)(239)))));
            this.btn備料單Unlock.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btn備料單Unlock.FlatAppearance.BorderSize = 0;
            this.btn備料單Unlock.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btn備料單Unlock.Font = new System.Drawing.Font("微軟正黑體", 12F, System.Drawing.FontStyle.Bold);
            this.btn備料單Unlock.ForeColor = System.Drawing.Color.White;
            this.btn備料單Unlock.Location = new System.Drawing.Point(991, 15);
            this.btn備料單Unlock.Name = "btn備料單Unlock";
            this.btn備料單Unlock.Size = new System.Drawing.Size(100, 38);
            this.btn備料單Unlock.TabIndex = 4;
            this.btn備料單Unlock.Text = "解鎖";
            this.btn備料單Unlock.UseVisualStyleBackColor = false;
            this.btn備料單Unlock.Click += new System.EventHandler(this.btn備料單Unlock_Click);
            // 
            // btn備料單返回
            // 
            this.btn備料單返回.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(100)))), ((int)(((byte)(100)))), ((int)(((byte)(100)))));
            this.btn備料單返回.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btn備料單返回.FlatAppearance.BorderSize = 0;
            this.btn備料單返回.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btn備料單返回.Font = new System.Drawing.Font("微軟正黑體", 12F, System.Drawing.FontStyle.Bold);
            this.btn備料單返回.ForeColor = System.Drawing.Color.White;
            this.btn備料單返回.Location = new System.Drawing.Point(1108, 15);
            this.btn備料單返回.Name = "btn備料單返回";
            this.btn備料單返回.Size = new System.Drawing.Size(123, 38);
            this.btn備料單返回.TabIndex = 6;
            this.btn備料單返回.Text = "返回主頁";
            this.btn備料單返回.UseVisualStyleBackColor = false;
            this.btn備料單返回.Click += new System.EventHandler(this.btn備料單返回_Click);
            // 
            // pnlTop
            // 
            this.pnlTop.BackColor = System.Drawing.Color.Transparent;
            this.pnlTop.Controls.Add(this.btn備料單存檔);
            this.pnlTop.Controls.Add(this.lbl料號);
            this.pnlTop.Controls.Add(this.txt備料單料號);
            this.pnlTop.Controls.Add(this.lbl數量);
            this.pnlTop.Controls.Add(this.txt備料單數量);
            this.pnlTop.Controls.Add(this.btn備料單匯入檔案);
            this.pnlTop.Controls.Add(this.btn備料單匯出);
            this.pnlTop.Controls.Add(this.btn備料單Unlock);
            this.pnlTop.Controls.Add(this.btn備料單返回);
            this.pnlTop.Dock = System.Windows.Forms.DockStyle.Top;
            this.pnlTop.Location = new System.Drawing.Point(0, 0);
            this.pnlTop.Name = "pnlTop";
            this.pnlTop.Padding = new System.Windows.Forms.Padding(9);
            this.pnlTop.Size = new System.Drawing.Size(1307, 70);
            this.pnlTop.TabIndex = 0;
            // 
            // btn備料單存檔
            // 
            this.btn備料單存檔.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(128)))), ((int)(((byte)(255)))), ((int)(((byte)(128)))));
            this.btn備料單存檔.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btn備料單存檔.FlatAppearance.BorderSize = 0;
            this.btn備料單存檔.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btn備料單存檔.Font = new System.Drawing.Font("微軟正黑體", 12F, System.Drawing.FontStyle.Bold);
            this.btn備料單存檔.ForeColor = System.Drawing.Color.DarkOliveGreen;
            this.btn備料單存檔.Location = new System.Drawing.Point(713, 13);
            this.btn備料單存檔.Name = "btn備料單存檔";
            this.btn備料單存檔.Size = new System.Drawing.Size(111, 38);
            this.btn備料單存檔.TabIndex = 7;
            this.btn備料單存檔.Text = "存檔";
            this.btn備料單存檔.UseVisualStyleBackColor = false;
            this.btn備料單存檔.Click += new System.EventHandler(this.btn備料單存檔_Click);
            // 
            // Form備料單匯入
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(12F, 25F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(245)))), ((int)(((byte)(245)))), ((int)(((byte)(245)))));
            this.ClientSize = new System.Drawing.Size(1307, 722);
            this.Controls.Add(this.dgv備料單);
            this.Controls.Add(this.pnlTop);
            this.Font = new System.Drawing.Font("微軟正黑體", 12F, System.Drawing.FontStyle.Bold);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.Name = "Form備料單匯入";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "備料單匯入";
            ((System.ComponentModel.ISupportInitialize)(this.dgv備料單)).EndInit();
            this.pnlTop.ResumeLayout(false);
            this.pnlTop.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.DataGridView dgv備料單;
        private System.Windows.Forms.Button btn備料單匯入檔案;
        private System.Windows.Forms.TextBox txt備料單料號;
        private System.Windows.Forms.TextBox txt備料單數量;
        private System.Windows.Forms.Button btn備料單Unlock;
        private System.Windows.Forms.Button btn備料單匯出;
    private System.Windows.Forms.Label lbl料號;
    private System.Windows.Forms.Label lbl數量;
        private System.Windows.Forms.Button btn備料單返回;
    private System.Windows.Forms.Panel pnlTop;
        private System.Windows.Forms.Button btn備料單存檔;
    }
}
