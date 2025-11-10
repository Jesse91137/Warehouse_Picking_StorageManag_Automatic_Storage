
namespace Automatic_Storage
{
    partial class PcbDC
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
            this.btn_pcbDC = new System.Windows.Forms.Button();
            this.txtPcbOld = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.txtPcbNew = new System.Windows.Forms.TextBox();
            this.SuspendLayout();
            // 
            // btn_pcbDC
            // 
            this.btn_pcbDC.Font = new System.Drawing.Font("微軟正黑體", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.btn_pcbDC.Location = new System.Drawing.Point(227, 125);
            this.btn_pcbDC.Name = "btn_pcbDC";
            this.btn_pcbDC.Size = new System.Drawing.Size(75, 33);
            this.btn_pcbDC.TabIndex = 10;
            this.btn_pcbDC.Text = "修改";
            this.btn_pcbDC.UseVisualStyleBackColor = true;
            this.btn_pcbDC.Click += new System.EventHandler(this.btn_pcbDC_Click);
            // 
            // txtPcbOld
            // 
            this.txtPcbOld.Font = new System.Drawing.Font("微軟正黑體", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.txtPcbOld.Location = new System.Drawing.Point(132, 15);
            this.txtPcbOld.Name = "txtPcbOld";
            this.txtPcbOld.ReadOnly = true;
            this.txtPcbOld.Size = new System.Drawing.Size(170, 33);
            this.txtPcbOld.TabIndex = 8;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("微軟正黑體", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label2.Location = new System.Drawing.Point(14, 82);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(84, 24);
            this.label2.TabIndex = 7;
            this.label2.Text = "異動PCB";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("微軟正黑體", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label1.Location = new System.Drawing.Point(14, 18);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(84, 24);
            this.label1.TabIndex = 6;
            this.label1.Text = "原始PCB";
            // 
            // txtPcbNew
            // 
            this.txtPcbNew.Font = new System.Drawing.Font("微軟正黑體", 14.25F);
            this.txtPcbNew.Location = new System.Drawing.Point(132, 73);
            this.txtPcbNew.Name = "txtPcbNew";
            this.txtPcbNew.Size = new System.Drawing.Size(170, 33);
            this.txtPcbNew.TabIndex = 11;
            // 
            // PcbDC
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(321, 176);
            this.Controls.Add(this.txtPcbNew);
            this.Controls.Add(this.btn_pcbDC);
            this.Controls.Add(this.txtPcbOld);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Name = "PcbDC";
            this.Text = "PCB DC";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.PcbDC_FormClosing);
            this.Load += new System.EventHandler(this.PcbDC_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btn_pcbDC;
        private System.Windows.Forms.TextBox txtPcbOld;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox txtPcbNew;
    }
}