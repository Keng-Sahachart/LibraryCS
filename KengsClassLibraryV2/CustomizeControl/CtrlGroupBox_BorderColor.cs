using System.Windows.Forms;
using System.Drawing;
using Microsoft.VisualBasic.CompilerServices;

namespace KengsLibraryCs
{
    public class CtrlGroupBox_BorderColor : GroupBox
    {
        private Color _bordercolor;

        public CtrlGroupBox_BorderColor() : base()
        {
            _bordercolor = Color.Black;
        }

        public Color BorderColor
        {
            get
            {
                return _bordercolor;
            }
            set
            {
                _bordercolor = value;
            }
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            {
                var withBlock = this;
                var sz = TextRenderer.MeasureText(withBlock.Text, withBlock.Font);
                var rect = withBlock.ClientRectangle;
                rect
.Height = Conversions.ToInteger(rect.Height - sz.Height / (double)2);
                rect.Y = Conversions.ToInteger(rect.Y + sz.Height / (double)2);
                ControlPaint.DrawBorder(e.Graphics, rect, _bordercolor, ButtonBorderStyle.Solid);
                var TextPosition = ClientRectangle;
                TextPosition
.Width = sz.Width;
                TextPosition.Height = sz.Height;
                TextPosition.X = TextPosition.X + 5;
                {
                    var withBlock1 = e;
                    withBlock1.Graphics.FillRectangle(new SolidBrush(BackColor), TextPosition);
                    withBlock1.Graphics.DrawString(Text, Font, new SolidBrush(ForeColor), TextPosition);
                }
            }
        }
    }
}
