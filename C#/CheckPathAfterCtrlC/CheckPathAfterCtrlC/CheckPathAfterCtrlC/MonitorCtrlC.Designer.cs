namespace CheckPathAfterCtrlC
{
    partial class MonitorCtrlC
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MonitorCtrlC));
            this.path = new System.Windows.Forms.Label();
            this.TextInClipborard = new System.Windows.Forms.TextBox();
            this.Valid = new System.Windows.Forms.Label();
            this.Invalid = new System.Windows.Forms.Label();
            this.Message = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // path
            // 
            this.path.AutoSize = true;
            this.path.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.path.Location = new System.Drawing.Point(12, 36);
            this.path.Name = "path";
            this.path.Size = new System.Drawing.Size(62, 13);
            this.path.TabIndex = 0;
            this.path.Text = "Ctrl + C -->>";
            // 
            // TextInClipborard
            // 
            this.TextInClipborard.Location = new System.Drawing.Point(80, 33);
            this.TextInClipborard.Name = "TextInClipborard";
            this.TextInClipborard.ReadOnly = true;
            this.TextInClipborard.Size = new System.Drawing.Size(500, 20);
            this.TextInClipborard.TabIndex = 1;
            this.TextInClipborard.Text = "Text in Clipboard";
            // 
            // Valid
            // 
            this.Valid.AutoSize = true;
            this.Valid.BackColor = System.Drawing.Color.Green;
            this.Valid.Location = new System.Drawing.Point(487, 11);
            this.Valid.Name = "Valid";
            this.Valid.Size = new System.Drawing.Size(30, 13);
            this.Valid.TabIndex = 2;
            this.Valid.Text = "Valid";
            // 
            // Invalid
            // 
            this.Invalid.AutoSize = true;
            this.Invalid.BackColor = System.Drawing.Color.Red;
            this.Invalid.Location = new System.Drawing.Point(542, 11);
            this.Invalid.Name = "Invalid";
            this.Invalid.Size = new System.Drawing.Size(38, 13);
            this.Invalid.TabIndex = 3;
            this.Invalid.Text = "Invalid";
            // 
            // Message
            // 
            this.Message.AutoSize = true;
            this.Message.Location = new System.Drawing.Point(77, 11);
            this.Message.Name = "Message";
            this.Message.Size = new System.Drawing.Size(0, 13);
            this.Message.TabIndex = 4;
            // 
            // MonitorCtrlC
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.Highlight;
            this.ClientSize = new System.Drawing.Size(603, 65);
            this.Controls.Add(this.Message);
            this.Controls.Add(this.Invalid);
            this.Controls.Add(this.Valid);
            this.Controls.Add(this.TextInClipborard);
            this.Controls.Add(this.path);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.Name = "MonitorCtrlC";
            this.Text = "Monitor Ctrl + C";
            this.TopMost = true;
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label path;
        private System.Windows.Forms.TextBox TextInClipborard;
        private System.Windows.Forms.Label Valid;
        private System.Windows.Forms.Label Invalid;
        private System.Windows.Forms.Label Message;
    }
}

