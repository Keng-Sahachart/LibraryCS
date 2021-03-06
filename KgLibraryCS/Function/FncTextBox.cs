﻿using System.Windows.Forms;
using System;

namespace kgLibraryCs
{
    /// <summary>
    /// 590114
    /// ต้องประกาศ ตัวแปร ก่อน
    /// ทำ Event handle 
    /// </summary>
    /// <remarks></remarks>
    public class FncTextBox
    {

        /// <summary>
        /// Set AddHandler ให้ TextBox
        /// </summary>
        public static void SetTextBoxForDragDrop(TextBox TextBoxDrop)
        {
            TextBoxDrop.AllowDrop = true;
            // TextBoxDrop 
            //TextBoxDrop.DragDrop += new DragEventHandler(this.TextBoxDrop_DragDrop);
            //TextBoxDrop.DragEnter += new DragEventHandler(this.TextBoxDrop_DragEnter);

            TextBoxDrop.DragDrop += new DragEventHandler(TextBoxDrop_DragDrop);
            TextBoxDrop.DragEnter += new DragEventHandler(TextBoxDrop_DragEnter);

        }

        public static void TextBoxDrop_DragDrop(object sender, DragEventArgs e) // Handles TextBoxDrop.DragDrop
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                string[] MyFiles;
                // Assign the files to an array.
                MyFiles = e.Data.GetData(DataFormats.FileDrop) as string[];
                // Display the file Name
                ((TextBox)sender).Text = MyFiles[0];
            }
        }
        public static void TextBoxDrop_DragEnter(object sender, DragEventArgs e) // Handles TextBoxDrop.DragEnter
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
                e.Effect = DragDropEffects.All;
        }
    }
}
