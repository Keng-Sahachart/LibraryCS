using System.Drawing;
using System.Windows.Forms;
using Microsoft.VisualBasic.CompilerServices;

namespace KengsLibraryCs
{
    /// <summary>
    /// Class ที่ พัฒนาจาก ToolTip ที่เป็น Text เพื่อให้แสดงเป็น Image ได้
    /// </summary>
    /// <remarks></remarks>
    public class CtrlImageToolTip : ToolTip
    {
        private Size PopupSize = new Size(640, 480);
        public CtrlImageToolTip() : base()
        {
            OwnerDraw = true; // Must be set otherwise will not draw the image properly
            IsBalloon = false;
        }

        private void ImageTip_Draw(object sender, DrawToolTipEventArgs e)
        {
            // Draws the image in the tooltip popup by reading in the tooltip text on the control that you want to have a popup for
            e.DrawBackground();
            e.DrawBorder();

            Image Image;
            Image = Image.FromFile(e.ToolTipText);
            // If Image.Width < 640 Then 'ถ้ารูปน้อยกว่า เกณที่กำหนด ให้ใช้ ขนาดรูปจริง
            // PopupSize.Width = Image.Width
            // PopupSize.Height = Image.Height
            // End If

            // Draws the image with the name stored in the tooltip text input for each control on the form
            // e.Graphics.DrawImage(My.Resources.ResourceManager.GetObject(e.ToolTipText), 0, 0)
            e.Graphics.DrawImage(Image, 0, 0, PopupSize.Width, PopupSize.Height); // Image.Width, Image.Height)
        }

        private void ImageTip_Popup(object sender, PopupEventArgs e)
        {
            // Creates the popup and sets its dimensions from the resource name
            Image Image;
            string ToolText;

            ToolText = GetToolTip(e.AssociatedControl);
            if (!string.IsNullOrEmpty(ToolText))
            {
                Image = Image.FromFile(ToolText); // My.Resources.ResourceManager.GetObject(ToolText) 'My.Resources.ResourceManager.GetObject(ToolText)

                if (Image.Width < 640)
                {
                    PopupSize.Width = Image.Width;
                    PopupSize.Height = Image.Height;
                }
                else
                {
                    float RatioW640 = Conversions.ToSingle(640 / (double)Image.Width);
                    PopupSize.Width = Conversions.ToInteger(Image.Width * RatioW640);
                    PopupSize.Height = Conversions.ToInteger(Image.Height * RatioW640);
                }

                if (Image != null)
                    e.ToolTipSize = new Size(PopupSize.Width, PopupSize.Height); // Size(Image.Width, Image.Height)
                else
                    e.Cancel = true;
            }
        }
    }
}



// #############################
// Original
// #############################
// Public Class ImageTip
// Inherits ToolTip
// Public Sub New()
// MyBase.New()
// Me.OwnerDraw = True 'Must be set otherwise will not draw the image properly
// Me.IsBalloon = False
// End Sub

// Private Sub ImageTip_Draw(ByVal sender As Object, ByVal e As DrawToolTipEventArgs) Handles Me.Draw
// 'Draws the image in the tooltip popup by reading in the tooltip text on the control that you want to have a popup for
// e.DrawBackground()
// e.DrawBorder()

// e.Graphics.DrawImage(My.Resources.ResourceManager.GetObject(e.ToolTipText), 0, 0) 'Draws the image with the name stored in the tooltip text input for each control on the form
// End Sub

// Private Sub ImageTip_Popup(ByVal sender As Object, ByVal e As PopupEventArgs) Handles Me.Popup
// 'Creates the popup and sets its dimensions from the resource name
// Dim Image As Image
// Dim ToolText As String

// ToolText = Me.GetToolTip(e.AssociatedControl)
// If ToolText <> "" Then
// Image = My.Resources.ResourceManager.GetObject(ToolText) 'My.Resources.ResourceManager.GetObject(ToolText)
// If Image IsNot Nothing Then
// e.ToolTipSize = New Size(Image.Width, Image.Height)
// Else
// e.Cancel = True
// End If
// End If
// End Sub
// End Class
// #########################################################################
