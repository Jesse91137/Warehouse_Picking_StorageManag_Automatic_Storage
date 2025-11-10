
namespace Automatic_Storage
{
    partial class Login
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Login));
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.txt_id = new System.Windows.Forms.TextBox();
            this.txt_psw = new System.Windows.Forms.TextBox();
            this.loginresult = new System.Windows.Forms.Label();
            this.button1 = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(30, 27);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(65, 12);
            this.label1.TabIndex = 0;
            this.label1.Text = "使用者代號";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(30, 68);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(65, 12);
            this.label2.TabIndex = 1;
            this.label2.Text = "使用者密碼";
            //this.label2.Click += new System.EventHandler(this.label2_Click);
            // 
            // txt_id
            // 
            this.txt_id.Font = new System.Drawing.Font("微軟正黑體", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.txt_id.Location = new System.Drawing.Point(101, 17);
            this.txt_id.Name = "txt_id";
            this.txt_id.Size = new System.Drawing.Size(131, 29);
            this.txt_id.TabIndex = 2;
            this.txt_id.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.textBox1_KeyPress);
            // 
            // txt_psw
            // 
            this.txt_psw.Font = new System.Drawing.Font("微軟正黑體", 12F, System.Drawing.FontStyle.Bold);
            this.txt_psw.Location = new System.Drawing.Point(101, 59);
            this.txt_psw.Name = "txt_psw";
            this.txt_psw.Size = new System.Drawing.Size(131, 29);
            this.txt_psw.TabIndex = 3;
            this.txt_psw.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txt_psw_KeyPress);
            // 
            // loginresult
            // 
            this.loginresult.AutoSize = true;
            this.loginresult.Font = new System.Drawing.Font("微軟正黑體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.loginresult.ForeColor = System.Drawing.Color.Red;
            this.loginresult.Location = new System.Drawing.Point(12, 98);
            this.loginresult.Name = "loginresult";
            this.loginresult.Size = new System.Drawing.Size(0, 20);
            this.loginresult.TabIndex = 1;
            //this.loginresult.Click += new System.EventHandler(this.label2_Click);
            // 
            // button1
            // 
            this.button1.BackColor = System.Drawing.Color.SeaGreen;
            this.button1.Font = new System.Drawing.Font("微軟正黑體", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.button1.ForeColor = System.Drawing.Color.Maroon;
            this.button1.Location = new System.Drawing.Point(177, 94);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(55, 35);
            this.button1.TabIndex = 4;
            this.button1.Text = "Exit";
            this.button1.UseVisualStyleBackColor = false;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // Login
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(246, 132);
            this.ControlBox = false;
            this.Controls.Add(this.button1);
            this.Controls.Add(this.txt_psw);
            this.Controls.Add(this.txt_id);
            this.Controls.Add(this.loginresult);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "Login";
            this.Text = "使用者登入";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox txt_id;
        private System.Windows.Forms.TextBox txt_psw;
        private System.Windows.Forms.Label loginresult;
        private System.Windows.Forms.Button button1;
    }
}