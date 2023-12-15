namespace FinancialAdministrator
{
    partial class financialAdministrator
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
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
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            filePreview = new RichTextBox();
            readFileButton = new Button();
            openFileDialog1 = new OpenFileDialog();
            generateXcelButton = new Button();
            progressBar1 = new ProgressBar();
            saveFileDialog1 = new SaveFileDialog();
            SuspendLayout();
            // 
            // filePreview
            // 
            filePreview.Location = new Point(69, 55);
            filePreview.Name = "filePreview";
            filePreview.Size = new Size(387, 279);
            filePreview.TabIndex = 0;
            filePreview.Text = "";
            // 
            // readFileButton
            // 
            readFileButton.Location = new Point(73, 351);
            readFileButton.Name = "readFileButton";
            readFileButton.Size = new Size(151, 29);
            readFileButton.TabIndex = 1;
            readFileButton.Text = "Open CSV bestand";
            readFileButton.UseVisualStyleBackColor = true;
            readFileButton.Click += readFileButton_Click;
            // 
            // openFileDialog1
            // 
            openFileDialog1.FileName = "openFileDialog1";
            // 
            // generateXcelButton
            // 
            generateXcelButton.Location = new Point(242, 351);
            generateXcelButton.Name = "generateXcelButton";
            generateXcelButton.Size = new Size(181, 29);
            generateXcelButton.TabIndex = 2;
            generateXcelButton.Text = "Genereer Excel bestand";
            generateXcelButton.UseVisualStyleBackColor = true;
            generateXcelButton.Click += generateXcelButton_Click;
            // 
            // progressBar1
            // 
            progressBar1.Location = new Point(69, 400);
            progressBar1.Name = "progressBar1";
            progressBar1.Size = new Size(387, 29);
            progressBar1.TabIndex = 3;
      
            // 
            // financialAdministrator
            // 
            AutoScaleDimensions = new SizeF(8F, 20F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(523, 450);
            Controls.Add(progressBar1);
            Controls.Add(generateXcelButton);
            Controls.Add(readFileButton);
            Controls.Add(filePreview);
            Name = "financialAdministrator";
            Text = "Financiele Administrator";
            ResumeLayout(false);
        }

        #endregion

        private RichTextBox filePreview;
        private Button readFileButton;
        private OpenFileDialog openFileDialog1;
        private Button generateXcelButton;
        private ProgressBar progressBar1;
        private SaveFileDialog saveFileDialog1;
    }
}
