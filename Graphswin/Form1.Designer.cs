﻿namespace Graphswin
{
    partial class Form1
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
            ucGraphs1 = new ThrusterTest.UserControls.ucGraphs();
            SuspendLayout();
            // 
            // ucGraphs1
            // 
            ucGraphs1.AutoSize = true;
            ucGraphs1.AutoSizeMode = AutoSizeMode.GrowAndShrink;
            ucGraphs1.Dock = DockStyle.Fill;
            ucGraphs1.Location = new Point(0, 0);
            ucGraphs1.Name = "ucGraphs1";
            ucGraphs1.Size = new Size(1073, 619);
            ucGraphs1.TabIndex = 0;
            // 
            // Form1
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(1073, 619);
            Controls.Add(ucGraphs1);
            Name = "Form1";
            Text = "Form1";
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private ThrusterTest.UserControls.ucGraphs ucGraphs1;
    }
}
