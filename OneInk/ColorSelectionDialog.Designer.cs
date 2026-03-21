/*
 *  Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license.
 *  ColorSelectionDialog.Designer.cs - Designer code for ColorSelectionDialog
 */

namespace OneInk
{
    partial class ColorSelectionDialog
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
            this.colorListView = new System.Windows.Forms.ListView();
            this.okButton = new System.Windows.Forms.Button();
            this.cancelButton = new System.Windows.Forms.Button();
            this.colorImageList = new System.Windows.Forms.ImageList();
            this.headerLabel = new System.Windows.Forms.Label();
            this.SuspendLayout();
            //
            // colorListView
            //
            this.colorListView.FullRowSelect = true;
            this.colorListView.SmallImageList = this.colorImageList;
            this.colorListView.Location = new System.Drawing.Point(12, 40);
            this.colorListView.MultiSelect = false;
            this.colorListView.Name = "colorListView";
            this.colorListView.Size = new System.Drawing.Size(340, 200);
            this.colorListView.TabIndex = 0;
            this.colorListView.UseCompatibleStateImageBehavior = false;
            this.colorListView.View = System.Windows.Forms.View.List;
            this.colorListView.SelectedIndexChanged += new System.EventHandler(this.colorListView_SelectedIndexChanged);
            this.colorListView.DoubleClick += new System.EventHandler(this.colorListView_DoubleClick);
            //
            // okButton
            //
            this.okButton.Location = new System.Drawing.Point(12, 250);
            this.okButton.Name = "okButton";
            this.okButton.Size = new System.Drawing.Size(160, 28);
            this.okButton.TabIndex = 1;
            this.okButton.Text = Strings.OkButton;
            this.okButton.UseVisualStyleBackColor = true;
            this.okButton.Click += new System.EventHandler(this.okButton_Click);
            //
            // cancelButton
            //
            this.cancelButton.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.cancelButton.Location = new System.Drawing.Point(192, 250);
            this.cancelButton.Name = "cancelButton";
            this.cancelButton.Size = new System.Drawing.Size(160, 28);
            this.cancelButton.TabIndex = 2;
            this.cancelButton.Text = Strings.CancelButton;
            this.cancelButton.UseVisualStyleBackColor = true;
            this.cancelButton.Click += new System.EventHandler(this.cancelButton_Click);
            //
            // colorImageList
            //
            this.colorImageList.ImageSize = new System.Drawing.Size(16, 16);
            //
            // headerLabel
            //
            this.headerLabel.AutoSize = true;
            this.headerLabel.Location = new System.Drawing.Point(12, 12);
            this.headerLabel.Name = "headerLabel";
            this.headerLabel.Size = new System.Drawing.Size(200, 13);
            this.headerLabel.TabIndex = 3;
            this.headerLabel.Text = Strings.DialogHeader;
            //
            // ColorSelectionDialog
            //
            this.AcceptButton = this.okButton;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.cancelButton;
            this.ClientSize = new System.Drawing.Size(364, 290);
            this.Controls.Add(this.headerLabel);
            this.Controls.Add(this.cancelButton);
            this.Controls.Add(this.okButton);
            this.Controls.Add(this.colorListView);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "ColorSelectionDialog";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = Strings.DialogTitle;
            this.ResumeLayout(false);
            this.PerformLayout();
        }

        #endregion

        private System.Windows.Forms.ListView colorListView;
        private System.Windows.Forms.Button okButton;
        private System.Windows.Forms.Button cancelButton;
        private System.Windows.Forms.ImageList colorImageList;
        private System.Windows.Forms.Label headerLabel;
    }
}
