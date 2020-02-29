namespace 麦克利米曲线
{
    partial class 用户主页
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
            this.labelSearch = new System.Windows.Forms.Label();
            this.labelSelect = new System.Windows.Forms.Label();
            this.labelUserC = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.ExitButton = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // labelSearch
            // 
            this.labelSearch.AutoSize = true;
            this.labelSearch.BackColor = System.Drawing.Color.Transparent;
            this.labelSearch.Cursor = System.Windows.Forms.Cursors.Hand;
            this.labelSearch.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.labelSearch.ForeColor = System.Drawing.SystemColors.ButtonHighlight;
            this.labelSearch.Location = new System.Drawing.Point(199, 268);
            this.labelSearch.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.labelSearch.Name = "labelSearch";
            this.labelSearch.Size = new System.Drawing.Size(78, 17);
            this.labelSearch.TabIndex = 2;
            this.labelSearch.Text = "数据可视化";
            this.labelSearch.Click += new System.EventHandler(this.labelSearch_Click);
            // 
            // labelSelect
            // 
            this.labelSelect.AutoSize = true;
            this.labelSelect.BackColor = System.Drawing.Color.Transparent;
            this.labelSelect.Cursor = System.Windows.Forms.Cursors.Hand;
            this.labelSelect.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.labelSelect.ForeColor = System.Drawing.SystemColors.ButtonHighlight;
            this.labelSelect.Location = new System.Drawing.Point(397, 268);
            this.labelSelect.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.labelSelect.Name = "labelSelect";
            this.labelSelect.Size = new System.Drawing.Size(64, 17);
            this.labelSelect.TabIndex = 3;
            this.labelSelect.Text = "海洋科普";
            this.labelSelect.Click += new System.EventHandler(this.labelSelect_Click);
            // 
            // labelUserC
            // 
            this.labelUserC.AutoSize = true;
            this.labelUserC.BackColor = System.Drawing.Color.Transparent;
            this.labelUserC.Cursor = System.Windows.Forms.Cursors.Hand;
            this.labelUserC.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.labelUserC.ForeColor = System.Drawing.SystemColors.ButtonHighlight;
            this.labelUserC.Location = new System.Drawing.Point(600, 268);
            this.labelUserC.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.labelUserC.Name = "labelUserC";
            this.labelUserC.Size = new System.Drawing.Size(64, 17);
            this.labelUserC.TabIndex = 4;
            this.labelUserC.Text = "用户中心";
            this.labelUserC.Click += new System.EventHandler(this.labelUserC_Click);
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.BackColor = System.Drawing.Color.Transparent;
            this.label4.Font = new System.Drawing.Font("Microsoft Sans Serif", 16F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label4.ForeColor = System.Drawing.SystemColors.ButtonHighlight;
            this.label4.Location = new System.Drawing.Point(309, 49);
            this.label4.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(212, 26);
            this.label4.TabIndex = 5;
            this.label4.Text = "To the Infinite Ocean";
            // 
            // ExitButton
            // 
            this.ExitButton.BackColor = System.Drawing.Color.Transparent;
            this.ExitButton.Cursor = System.Windows.Forms.Cursors.Hand;
            this.ExitButton.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.ExitButton.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.ExitButton.ForeColor = System.Drawing.Color.DarkGray;
            this.ExitButton.Location = new System.Drawing.Point(355, 410);
            this.ExitButton.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.ExitButton.Name = "ExitButton";
            this.ExitButton.Size = new System.Drawing.Size(160, 49);
            this.ExitButton.TabIndex = 10;
            this.ExitButton.Text = "退出登录";
            this.ExitButton.UseVisualStyleBackColor = false;
            this.ExitButton.Click += new System.EventHandler(this.ExitButton_Click);
            // 
            // 用户主页
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackgroundImage = global::麦克利米曲线.Properties.Resources.啦啦;
            this.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.ClientSize = new System.Drawing.Size(858, 478);
            this.Controls.Add(this.ExitButton);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.labelUserC);
            this.Controls.Add(this.labelSelect);
            this.Controls.Add(this.labelSearch);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.Name = "用户主页";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "求职者主页";
            this.Load += new System.EventHandler(this.用户主页_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label labelSearch;
        private System.Windows.Forms.Label labelSelect;
        private System.Windows.Forms.Label labelUserC;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Button ExitButton;
    }
}