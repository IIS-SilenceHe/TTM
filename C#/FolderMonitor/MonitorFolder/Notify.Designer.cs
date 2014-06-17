namespace MonitorFolder
{
    partial class Notify
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
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Notify));
            this.showMessage = new System.Windows.Forms.Label();
            this.OK = new System.Windows.Forms.Button();
            this.Cancle = new System.Windows.Forms.Button();
            this.LinkToPath = new System.Windows.Forms.LinkLabel();
            this.message2 = new System.Windows.Forms.Label();
            this.UpdatedTime = new System.Windows.Forms.Label();
            this.currentTime = new System.Windows.Forms.Label();
            this.timer1 = new System.Windows.Forms.Timer(this.components);
            this.SuspendLayout();
            // 
            // showMessage
            // 
            this.showMessage.AutoSize = true;
            this.showMessage.Location = new System.Drawing.Point(32, 30);
            this.showMessage.Name = "showMessage";
            this.showMessage.Size = new System.Drawing.Size(75, 13);
            this.showMessage.TabIndex = 0;
            this.showMessage.Text = "showMessage";
            // 
            // OK
            // 
            this.OK.Location = new System.Drawing.Point(149, 157);
            this.OK.Name = "OK";
            this.OK.Size = new System.Drawing.Size(75, 23);
            this.OK.TabIndex = 1;
            this.OK.Text = "Continue";
            this.OK.UseVisualStyleBackColor = true;
            this.OK.Click += new System.EventHandler(this.OK_Click);
            this.OK.GotFocus += new System.EventHandler(this.OK_GotFocus);
            // 
            // Cancle
            // 
            this.Cancle.Location = new System.Drawing.Point(314, 157);
            this.Cancle.Name = "Cancle";
            this.Cancle.Size = new System.Drawing.Size(75, 23);
            this.Cancle.TabIndex = 2;
            this.Cancle.Text = "Stop";
            this.Cancle.UseVisualStyleBackColor = true;
            this.Cancle.Click += new System.EventHandler(this.Cancle_Click);
            // 
            // LinkToPath
            // 
            this.LinkToPath.AutoSize = true;
            this.LinkToPath.Location = new System.Drawing.Point(90, 70);
            this.LinkToPath.Name = "LinkToPath";
            this.LinkToPath.Size = new System.Drawing.Size(55, 13);
            this.LinkToPath.TabIndex = 3;
            this.LinkToPath.TabStop = true;
            this.LinkToPath.Text = "linkLabel1";
            this.LinkToPath.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.LinkToPath_LinkClicked);
            // 
            // message2
            // 
            this.message2.AutoSize = true;
            this.message2.Location = new System.Drawing.Point(32, 70);
            this.message2.Name = "message2";
            this.message2.Size = new System.Drawing.Size(52, 13);
            this.message2.TabIndex = 4;
            this.message2.Text = "See Link:";
            // 
            // UpdatedTime
            // 
            this.UpdatedTime.AutoSize = true;
            this.UpdatedTime.Location = new System.Drawing.Point(32, 110);
            this.UpdatedTime.Name = "UpdatedTime";
            this.UpdatedTime.Size = new System.Drawing.Size(77, 13);
            this.UpdatedTime.TabIndex = 5;
            this.UpdatedTime.Text = "Updated Time:";
            // 
            // currentTime
            // 
            this.currentTime.AutoSize = true;
            this.currentTime.Location = new System.Drawing.Point(255, 110);
            this.currentTime.Name = "currentTime";
            this.currentTime.Size = new System.Drawing.Size(63, 13);
            this.currentTime.TabIndex = 6;
            this.currentTime.Text = "currentTime";
            // 
            // timer1
            // 
            this.timer1.Tick += new System.EventHandler(this.timer1_Tick);
            // 
            // Notify
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.ActiveBorder;
            this.ClientSize = new System.Drawing.Size(538, 205);
            this.Controls.Add(this.currentTime);
            this.Controls.Add(this.UpdatedTime);
            this.Controls.Add(this.message2);
            this.Controls.Add(this.LinkToPath);
            this.Controls.Add(this.Cancle);
            this.Controls.Add(this.OK);
            this.Controls.Add(this.showMessage);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.Name = "Notify";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Notify";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label showMessage;
        private System.Windows.Forms.Button OK;
        private System.Windows.Forms.Button Cancle;
        private System.Windows.Forms.LinkLabel LinkToPath;
        private System.Windows.Forms.Label message2;
        private System.Windows.Forms.Label UpdatedTime;
        private System.Windows.Forms.Label currentTime;
        private System.Windows.Forms.Timer timer1;
    }
}