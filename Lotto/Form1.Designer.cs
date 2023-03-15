namespace LottoNumberGenerator
{
    partial class Form1
    {
        /// <summary>
        /// 필수 디자이너 변수입니다.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 사용 중인 모든 리소스를 정리합니다.
        /// </summary>
        /// <param name="disposing">관리되는 리소스를 해제해야 하는지 여부입니다.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form 디자이너에서 생성한 코드

        /// <summary>
        /// 디자이너 지원에 필요한 메서드입니다. 
        /// 이 메서드의 내용을 코드 편집기로 수정하지 마세요.
        /// </summary>
        private void InitializeComponent()
        {
            this.btnOpenExcel = new System.Windows.Forms.Button();
            this.lblFilePath = new System.Windows.Forms.Label();
            this.lblStatus = new System.Windows.Forms.Label();
            this.btnGenerateNumbers = new System.Windows.Forms.Button();
            this.listBoxNumbers = new System.Windows.Forms.ListBox();
            this.SuspendLayout();
            // 
            // btnOpenExcel
            // 
            this.btnOpenExcel.Location = new System.Drawing.Point(12, 12);
            this.btnOpenExcel.Name = "btnOpenExcel";
            this.btnOpenExcel.Size = new System.Drawing.Size(103, 23);
            this.btnOpenExcel.TabIndex = 0;
            this.btnOpenExcel.Text = "Open Excel File";
            this.btnOpenExcel.UseVisualStyleBackColor = true;
            this.btnOpenExcel.Click += new System.EventHandler(this.btnOpenExcel_Click);
            // 
            // lblFilePath
            // 
            this.lblFilePath.AutoSize = true;
            this.lblFilePath.Location = new System.Drawing.Point(121, 17);
            this.lblFilePath.Name = "lblFilePath";
            this.lblFilePath.Size = new System.Drawing.Size(0, 12);
            this.lblFilePath.TabIndex = 1;
            // 
            // lblStatus
            // 
            this.lblStatus.AutoSize = true;
            this.lblStatus.Location = new System.Drawing.Point(12, 51);
            this.lblStatus.Name = "lblStatus";
            this.lblStatus.Size = new System.Drawing.Size(0, 12);
            this.lblStatus.TabIndex = 2;
            // 
            // btnGenerateNumbers
            // 
            this.btnGenerateNumbers.Location = new System.Drawing.Point(12, 80);
            this.btnGenerateNumbers.Name = "btnGenerateNumbers";
            this.btnGenerateNumbers.Size = new System.Drawing.Size(117, 23);
            this.btnGenerateNumbers.TabIndex = 3;
            this.btnGenerateNumbers.Text = "Generate Numbers";
            this.btnGenerateNumbers.UseVisualStyleBackColor = true;
            this.btnGenerateNumbers.Click += new System.EventHandler(this.btnGenerateNumbers_Click);
            // 
            // listBoxNumbers
            // 
            this.listBoxNumbers.FormattingEnabled = true;
            this.listBoxNumbers.ItemHeight = 12;
            this.listBoxNumbers.Location = new System.Drawing.Point(135, 80);
            this.listBoxNumbers.Name = "listBoxNumbers";
            this.listBoxNumbers.Size = new System.Drawing.Size(295, 184);
            this.listBoxNumbers.TabIndex = 4;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(442, 274);
            this.Controls.Add(this.listBoxNumbers);
            this.Controls.Add(this.btnGenerateNumbers);
            this.Controls.Add(this.lblStatus);
            this.Controls.Add(this.lblFilePath);
            this.Controls.Add(this.btnOpenExcel);
            this.Name = "Form1";
            this.Text = "Lotto Number Generator";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnOpenExcel;
        private System.Windows.Forms.Label lblFilePath;
        private System.Windows.Forms.Label lblStatus;
        private System.Windows.Forms.Button btnGenerateNumbers;
        private System.Windows.Forms.ListBox listBoxNumbers;
    }
}
