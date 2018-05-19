namespace ExCell_Art
{
    partial class MainForm
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MainForm));
            this.InputName = new System.Windows.Forms.Label();
            this.OutputPathLabel = new System.Windows.Forms.Label();
            this.pathButton = new System.Windows.Forms.Button();
            this.Start = new System.Windows.Forms.Button();
            this.Cancel = new System.Windows.Forms.Button();
            this.progressBar = new System.Windows.Forms.ProgressBar();
            this.ImageButton = new System.Windows.Forms.Button();
            this.openFileDialog = new System.Windows.Forms.OpenFileDialog();
            this.saveFileDialog = new System.Windows.Forms.SaveFileDialog();
            this.TimeRemainingLabel = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // InputName
            // 
            this.InputName.AutoSize = true;
            this.InputName.Location = new System.Drawing.Point(175, 22);
            this.InputName.Name = "InputName";
            this.InputName.Size = new System.Drawing.Size(37, 13);
            this.InputName.TabIndex = 1;
            this.InputName.Text = "(none)";
            // 
            // OutputPathLabel
            // 
            this.OutputPathLabel.AutoSize = true;
            this.OutputPathLabel.Location = new System.Drawing.Point(175, 76);
            this.OutputPathLabel.Name = "OutputPathLabel";
            this.OutputPathLabel.Size = new System.Drawing.Size(99, 13);
            this.OutputPathLabel.TabIndex = 2;
            this.OutputPathLabel.Text = "No Output selected";
            // 
            // pathButton
            // 
            this.pathButton.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
            this.pathButton.Location = new System.Drawing.Point(12, 71);
            this.pathButton.Name = "pathButton";
            this.pathButton.Size = new System.Drawing.Size(123, 23);
            this.pathButton.TabIndex = 3;
            this.pathButton.Text = "Choose file path";
            this.pathButton.UseVisualStyleBackColor = true;
            this.pathButton.Click += new System.EventHandler(this.pathButton_Click);
            // 
            // Start
            // 
            this.Start.Location = new System.Drawing.Point(12, 271);
            this.Start.Name = "Start";
            this.Start.Size = new System.Drawing.Size(75, 23);
            this.Start.TabIndex = 4;
            this.Start.Text = "Start";
            this.Start.UseVisualStyleBackColor = true;
            this.Start.Click += new System.EventHandler(this.Start_Click);
            // 
            // Cancel
            // 
            this.Cancel.Enabled = false;
            this.Cancel.Location = new System.Drawing.Point(127, 271);
            this.Cancel.Name = "Cancel";
            this.Cancel.Size = new System.Drawing.Size(75, 23);
            this.Cancel.TabIndex = 5;
            this.Cancel.Text = "Cancel";
            this.Cancel.UseVisualStyleBackColor = true;
            this.Cancel.Click += new System.EventHandler(this.Cancel_Click);
            // 
            // progressBar
            // 
            this.progressBar.Location = new System.Drawing.Point(12, 231);
            this.progressBar.Name = "progressBar";
            this.progressBar.Size = new System.Drawing.Size(402, 23);
            this.progressBar.Step = 1;
            this.progressBar.Style = System.Windows.Forms.ProgressBarStyle.Continuous;
            this.progressBar.TabIndex = 6;
            // 
            // ImageButton
            // 
            this.ImageButton.Location = new System.Drawing.Point(12, 17);
            this.ImageButton.Name = "ImageButton";
            this.ImageButton.Size = new System.Drawing.Size(123, 23);
            this.ImageButton.TabIndex = 1;
            this.ImageButton.Text = "Choose source image";
            this.ImageButton.UseVisualStyleBackColor = true;
            this.ImageButton.Click += new System.EventHandler(this.ImageButton_Click);
            // 
            // openFileDialog
            // 
            this.openFileDialog.FileName = "openFileDialog";
            // 
            // TimeRemainingLabel
            // 
            this.TimeRemainingLabel.AutoSize = true;
            this.TimeRemainingLabel.Location = new System.Drawing.Point(251, 281);
            this.TimeRemainingLabel.Name = "TimeRemainingLabel";
            this.TimeRemainingLabel.Size = new System.Drawing.Size(81, 13);
            this.TimeRemainingLabel.TabIndex = 7;
            this.TimeRemainingLabel.Text = "Time remaining:";
            this.TimeRemainingLabel.Click += new System.EventHandler(this.TimeRemainingLabel_Click);
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(498, 350);
            this.Controls.Add(this.TimeRemainingLabel);
            this.Controls.Add(this.ImageButton);
            this.Controls.Add(this.progressBar);
            this.Controls.Add(this.Cancel);
            this.Controls.Add(this.Start);
            this.Controls.Add(this.pathButton);
            this.Controls.Add(this.OutputPathLabel);
            this.Controls.Add(this.InputName);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "MainForm";
            this.Text = "ExCell Art Generator";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.Label InputName;
        private System.Windows.Forms.Label OutputPathLabel;
        private System.Windows.Forms.Button pathButton;
        private System.Windows.Forms.Button Start;
        private System.Windows.Forms.Button Cancel;
        private System.Windows.Forms.ProgressBar progressBar;
        private System.Windows.Forms.Button ImageButton;
        private System.Windows.Forms.OpenFileDialog openFileDialog;
        private System.Windows.Forms.SaveFileDialog saveFileDialog;
        private System.Windows.Forms.Label TimeRemainingLabel;
    }
}

