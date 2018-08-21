namespace OneWinMonitor
{
    partial class showResultsForm
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
            this.resultDataGrid = new System.Windows.Forms.DataGridView();
            this.totalLabel = new System.Windows.Forms.Label();
            this.selectedPeriod = new System.Windows.Forms.Label();
            this.InToExcelButton = new System.Windows.Forms.Button();
            this.saveFileDialog1 = new System.Windows.Forms.SaveFileDialog();
            this.button1 = new System.Windows.Forms.Button();
            this.saveFileDialog2 = new System.Windows.Forms.SaveFileDialog();
            ((System.ComponentModel.ISupportInitialize)(this.resultDataGrid)).BeginInit();
            this.SuspendLayout();
            // 
            // resultDataGrid
            // 
            this.resultDataGrid.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.resultDataGrid.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells;
            this.resultDataGrid.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells;
            this.resultDataGrid.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.resultDataGrid.Location = new System.Drawing.Point(12, 29);
            this.resultDataGrid.Name = "resultDataGrid";
            this.resultDataGrid.RowHeadersWidth = 120;
            this.resultDataGrid.Size = new System.Drawing.Size(1229, 716);
            this.resultDataGrid.TabIndex = 0;
            // 
            // totalLabel
            // 
            this.totalLabel.AutoSize = true;
            this.totalLabel.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.totalLabel.Location = new System.Drawing.Point(12, 6);
            this.totalLabel.Name = "totalLabel";
            this.totalLabel.Size = new System.Drawing.Size(51, 20);
            this.totalLabel.TabIndex = 1;
            this.totalLabel.Text = "label1";
            // 
            // selectedPeriod
            // 
            this.selectedPeriod.AutoSize = true;
            this.selectedPeriod.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.selectedPeriod.Location = new System.Drawing.Point(415, 6);
            this.selectedPeriod.Name = "selectedPeriod";
            this.selectedPeriod.Size = new System.Drawing.Size(51, 20);
            this.selectedPeriod.TabIndex = 2;
            this.selectedPeriod.Text = "label1";
            // 
            // InToExcelButton
            // 
            this.InToExcelButton.Location = new System.Drawing.Point(899, 3);
            this.InToExcelButton.Name = "InToExcelButton";
            this.InToExcelButton.Size = new System.Drawing.Size(75, 23);
            this.InToExcelButton.TabIndex = 3;
            this.InToExcelButton.Text = "В excel";
            this.InToExcelButton.UseVisualStyleBackColor = true;
            this.InToExcelButton.Click += new System.EventHandler(this.InToExcelButton_Click);
            // 
            // saveFileDialog1
            // 
            this.saveFileDialog1.Filter = "Excel files (*.xlsx)|*.xlsx";
            this.saveFileDialog1.FileOk += new System.ComponentModel.CancelEventHandler(this.saveFileDialog1_FileOk);
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(726, 3);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(75, 23);
            this.button1.TabIndex = 4;
            this.button1.Text = "Статистика";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // saveFileDialog2
            // 
            this.saveFileDialog2.Filter = "Excel files (*.xlsx)|*.xlsx";
            this.saveFileDialog2.FileOk += new System.ComponentModel.CancelEventHandler(this.saveFileDialog2_FileOk);
            // 
            // showResultsForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1246, 757);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.InToExcelButton);
            this.Controls.Add(this.selectedPeriod);
            this.Controls.Add(this.totalLabel);
            this.Controls.Add(this.resultDataGrid);
            this.Name = "showResultsForm";
            this.Text = "showResultsForm";
            this.Load += new System.EventHandler(this.showResultsForm_Load);
            ((System.ComponentModel.ISupportInitialize)(this.resultDataGrid)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.DataGridView resultDataGrid;
        private System.Windows.Forms.Label totalLabel;
        private System.Windows.Forms.Label selectedPeriod;
        private System.Windows.Forms.Button InToExcelButton;
        private System.Windows.Forms.SaveFileDialog saveFileDialog1;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.SaveFileDialog saveFileDialog2;
    }
}