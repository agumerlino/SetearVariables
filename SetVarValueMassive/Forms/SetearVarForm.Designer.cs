namespace SetVarValueMassive
{
    partial class SetearVarForm
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
            this.btnSelectFile = new System.Windows.Forms.Button();
            this.btnSetVar = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.button3 = new System.Windows.Forms.Button();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.lblEstado = new System.Windows.Forms.Label();
            this.rdFiles = new System.Windows.Forms.RadioButton();
            this.rdFolder = new System.Windows.Forms.RadioButton();
            this.SuspendLayout();
            // 
            // btnSelectFile
            // 
            this.btnSelectFile.Location = new System.Drawing.Point(165, 134);
            this.btnSelectFile.Name = "btnSelectFile";
            this.btnSelectFile.Size = new System.Drawing.Size(140, 47);
            this.btnSelectFile.TabIndex = 0;
            this.btnSelectFile.Text = "Select excel file";
            this.btnSelectFile.UseVisualStyleBackColor = true;
            this.btnSelectFile.Click += new System.EventHandler(this.button1_Click);
            // 
            // btnSetVar
            // 
            this.btnSetVar.Location = new System.Drawing.Point(31, 235);
            this.btnSetVar.Name = "btnSetVar";
            this.btnSetVar.Size = new System.Drawing.Size(426, 56);
            this.btnSetVar.TabIndex = 1;
            this.btnSetVar.Text = "Set variable value";
            this.btnSetVar.UseVisualStyleBackColor = true;
            this.btnSetVar.Click += new System.EventHandler(this.btnSetVar_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(140, 118);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(196, 13);
            this.label1.TabIndex = 2;
            this.label1.Text = "Select an Excel document to continue...";
            // 
            // button3
            // 
            this.button3.Location = new System.Drawing.Point(352, 318);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(105, 23);
            this.button3.TabIndex = 3;
            this.button3.Text = "Exit";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.button3_Click);
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // lblEstado
            // 
            this.lblEstado.AutoSize = true;
            this.lblEstado.Location = new System.Drawing.Point(207, 194);
            this.lblEstado.Name = "lblEstado";
            this.lblEstado.Size = new System.Drawing.Size(0, 13);
            this.lblEstado.TabIndex = 4;
            // 
            // rdFiles
            // 
            this.rdFiles.AutoSize = true;
            this.rdFiles.Location = new System.Drawing.Point(113, 36);
            this.rdFiles.Name = "rdFiles";
            this.rdFiles.Size = new System.Drawing.Size(46, 17);
            this.rdFiles.TabIndex = 5;
            this.rdFiles.TabStop = true;
            this.rdFiles.Text = "Files";
            this.rdFiles.UseVisualStyleBackColor = true;
            this.rdFiles.CheckedChanged += new System.EventHandler(this.rdFiles_CheckedChanged);
            // 
            // rdFolder
            // 
            this.rdFolder.AutoSize = true;
            this.rdFolder.Location = new System.Drawing.Point(314, 36);
            this.rdFolder.Name = "rdFolder";
            this.rdFolder.Size = new System.Drawing.Size(59, 17);
            this.rdFolder.TabIndex = 6;
            this.rdFolder.TabStop = true;
            this.rdFolder.Text = "Folders";
            this.rdFolder.UseVisualStyleBackColor = true;
            this.rdFolder.CheckedChanged += new System.EventHandler(this.rdFolder_CheckedChanged);
            // 
            // SetearVarForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(488, 370);
            this.Controls.Add(this.rdFolder);
            this.Controls.Add(this.rdFiles);
            this.Controls.Add(this.lblEstado);
            this.Controls.Add(this.button3);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.btnSetVar);
            this.Controls.Add(this.btnSelectFile);
            this.Name = "SetearVarForm";
            this.Text = "Setting variables";
            this.TopMost = true;
            this.Load += new System.EventHandler(this.SetearVarForm_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnSelectFile;
        private System.Windows.Forms.Button btnSetVar;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.Label lblEstado;
        private System.Windows.Forms.RadioButton rdFiles;
        private System.Windows.Forms.RadioButton rdFolder;
    }
}