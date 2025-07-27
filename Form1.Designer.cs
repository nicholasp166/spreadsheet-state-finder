namespace cswf;
using System;
using System.Windows.Forms;


partial class Form1
{
    private System.Windows.Forms.Button uploadButton;
    private System.Windows.Forms.OpenFileDialog openFileDialog;
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
        this.components = new System.ComponentModel.Container();
        this.uploadButton = new System.Windows.Forms.Button();
        this.openFileDialog = new System.Windows.Forms.OpenFileDialog();

        this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
        this.ClientSize = new System.Drawing.Size(800, 450);
        this.Text = "Form1";

        // 
        // uploadButton
        // 
        this.uploadButton.Location = new System.Drawing.Point(350, 200);
        this.uploadButton.Name = "uploadButton";
        this.uploadButton.Size = new System.Drawing.Size(100, 30);
        this.uploadButton.Text = "Upload File";
        this.uploadButton.UseVisualStyleBackColor = true;
        this.uploadButton.Click += new System.EventHandler(this.uploadButton_Click);

        // 
        // openFileDialog
        // 
        this.openFileDialog.Filter = "CSV files (*.csv)|*.csv|Excel files (*.xlsx)|*.xlsx";
        this.openFileDialog.Title = "Select a CSV or Excel file";

        // 
        // Form1
        // 
        this.Controls.Add(this.uploadButton);
    }

    #endregion
}
