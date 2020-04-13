using System.Drawing;

namespace KengsLibraryCs
{
    /// <summary>
    /// Class ที่เก็บสีของ Excel 56 สี ที่เป็น Parameter Interior Color ของ Excel 2003 ขึ้นไป
    /// โดยเฉพาะ Excel 2003 ที่จะมีสีให้ใช้แค่ 56 สี
    /// </summary>
    /// <remarks></remarks>
    public class ClsExcelInteriorColor
    {

        struct PropertyInteriorColor
        {
            public Color ColorRGB;
            public string Name;
            public int ColorIndex;
        }

        public  PropertyInteriorColor[] ExcelInteriorColors ;// = new  PropertyInteriorColor[57];

        ClsExcelInteriorColor()
        {
            ExcelInteriorColors = new PropertyInteriorColor[57];
            // InteriorColors(1).ColorRGB = Color.FromArgb(0, 0, 0) : InteriorColors(1).Name = ""
            ExcelInteriorColors[1].ColorRGB = Color.FromArgb(0, 0, 0); ExcelInteriorColors[1].Name = "Black"; ExcelInteriorColors[1].ColorIndex = 1;
            ExcelInteriorColors[2].ColorRGB = Color.FromArgb(255, 255, 255); ExcelInteriorColors[2].Name = "White"; ExcelInteriorColors[2].ColorIndex = 2;
            ExcelInteriorColors[3].ColorRGB = Color.FromArgb(255, 0, 0); ExcelInteriorColors[3].Name = "Red"; ExcelInteriorColors[3].ColorIndex = 3;
            ExcelInteriorColors[4].ColorRGB = Color.FromArgb(0, 255, 0); ExcelInteriorColors[4].Name = "Bright Green"; ExcelInteriorColors[4].ColorIndex = 4;
            ExcelInteriorColors[5].ColorRGB = Color.FromArgb(0, 0, 255); ExcelInteriorColors[5].Name = "Blue"; ExcelInteriorColors[5].ColorIndex = 5;
            ExcelInteriorColors[6].ColorRGB = Color.FromArgb(255, 255, 0); ExcelInteriorColors[6].Name = "Yellow"; ExcelInteriorColors[6].ColorIndex = 6;
            ExcelInteriorColors[7].ColorRGB = Color.FromArgb(255, 0, 255); ExcelInteriorColors[7].Name = "Pink"; ExcelInteriorColors[7].ColorIndex = 7;
            ExcelInteriorColors[8].ColorRGB = Color.FromArgb(0, 255, 255); ExcelInteriorColors[8].Name = "Turquoise"; ExcelInteriorColors[8].ColorIndex = 8;
            ExcelInteriorColors[9].ColorRGB = Color.FromArgb(128, 0, 0); ExcelInteriorColors[9].Name = "Dark Red"; ExcelInteriorColors[9].ColorIndex = 9;
            ExcelInteriorColors[10].ColorRGB = Color.FromArgb(0, 128, 0); ExcelInteriorColors[10].Name = "Green"; ExcelInteriorColors[10].ColorIndex = 10;
            ExcelInteriorColors[11].ColorRGB = Color.FromArgb(0, 0, 128); ExcelInteriorColors[11].Name = "Dark Blue"; ExcelInteriorColors[11].ColorIndex = 11;
            ExcelInteriorColors[12].ColorRGB = Color.FromArgb(128, 128, 0); ExcelInteriorColors[12].Name = "Dark Yellow"; ExcelInteriorColors[12].ColorIndex = 12;
            ExcelInteriorColors[13].ColorRGB = Color.FromArgb(128, 0, 128); ExcelInteriorColors[13].Name = "Violet"; ExcelInteriorColors[13].ColorIndex = 13;
            ExcelInteriorColors[14].ColorRGB = Color.FromArgb(0, 128, 128); ExcelInteriorColors[14].Name = "Teal"; ExcelInteriorColors[14].ColorIndex = 14;
            ExcelInteriorColors[15].ColorRGB = Color.FromArgb(192, 192, 192); ExcelInteriorColors[15].Name = "Gray-25%"; ExcelInteriorColors[15].ColorIndex = 15;
            ExcelInteriorColors[16].ColorRGB = Color.FromArgb(128, 128, 128); ExcelInteriorColors[16].Name = "Gray-50%"; ExcelInteriorColors[16].ColorIndex = 16;
            ExcelInteriorColors[17].ColorRGB = Color.FromArgb(153, 153, 255); ExcelInteriorColors[17].Name = "Periwinkle"; ExcelInteriorColors[17].ColorIndex = 17;
            ExcelInteriorColors[18].ColorRGB = Color.FromArgb(153, 51, 102); ExcelInteriorColors[18].Name = "Plum+"; ExcelInteriorColors[18].ColorIndex = 18;
            ExcelInteriorColors[19].ColorRGB = Color.FromArgb(255, 255, 204); ExcelInteriorColors[19].Name = "Ivory"; ExcelInteriorColors[19].ColorIndex = 19;
            ExcelInteriorColors[20].ColorRGB = Color.FromArgb(204, 255, 255); ExcelInteriorColors[20].Name = "Lite Turquoise"; ExcelInteriorColors[20].ColorIndex = 20;
            ExcelInteriorColors[21].ColorRGB = Color.FromArgb(102, 0, 102); ExcelInteriorColors[21].Name = "Dark Purple"; ExcelInteriorColors[21].ColorIndex = 21;
            ExcelInteriorColors[22].ColorRGB = Color.FromArgb(255, 128, 128); ExcelInteriorColors[22].Name = "Coral"; ExcelInteriorColors[22].ColorIndex = 22;
            ExcelInteriorColors[23].ColorRGB = Color.FromArgb(0, 102, 204); ExcelInteriorColors[23].Name = "Ocean Blue"; ExcelInteriorColors[23].ColorIndex = 23;
            ExcelInteriorColors[24].ColorRGB = Color.FromArgb(204, 204, 255); ExcelInteriorColors[24].Name = "Ice Blue"; ExcelInteriorColors[24].ColorIndex = 24;
            ExcelInteriorColors[25].ColorRGB = Color.FromArgb(0, 0, 128); ExcelInteriorColors[25].Name = "Dark Blue+"; ExcelInteriorColors[25].ColorIndex = 25;
            ExcelInteriorColors[26].ColorRGB = Color.FromArgb(255, 0, 255); ExcelInteriorColors[26].Name = "Pink+"; ExcelInteriorColors[26].ColorIndex = 26;
            ExcelInteriorColors[27].ColorRGB = Color.FromArgb(255, 255, 0); ExcelInteriorColors[27].Name = "Yellow+"; ExcelInteriorColors[27].ColorIndex = 27;
            ExcelInteriorColors[28].ColorRGB = Color.FromArgb(0, 255, 255); ExcelInteriorColors[28].Name = "Turquoise+"; ExcelInteriorColors[28].ColorIndex = 28;
            ExcelInteriorColors[29].ColorRGB = Color.FromArgb(128, 0, 128); ExcelInteriorColors[29].Name = "Violet+"; ExcelInteriorColors[29].ColorIndex = 29;
            ExcelInteriorColors[30].ColorRGB = Color.FromArgb(128, 0, 0); ExcelInteriorColors[30].Name = "Dark Red+"; ExcelInteriorColors[30].ColorIndex = 30;
            ExcelInteriorColors[31].ColorRGB = Color.FromArgb(0, 128, 128); ExcelInteriorColors[31].Name = "Teal+"; ExcelInteriorColors[31].ColorIndex = 31;
            ExcelInteriorColors[32].ColorRGB = Color.FromArgb(0, 0, 255); ExcelInteriorColors[32].Name = "Blue+"; ExcelInteriorColors[32].ColorIndex = 32;
            ExcelInteriorColors[33].ColorRGB = Color.FromArgb(0, 204, 255); ExcelInteriorColors[33].Name = "Sky Blue"; ExcelInteriorColors[33].ColorIndex = 33;
            ExcelInteriorColors[34].ColorRGB = Color.FromArgb(204, 255, 255); ExcelInteriorColors[34].Name = "Light Turquoise"; ExcelInteriorColors[34].ColorIndex = 34;
            ExcelInteriorColors[35].ColorRGB = Color.FromArgb(204, 255, 204); ExcelInteriorColors[35].Name = "Light Green"; ExcelInteriorColors[35].ColorIndex = 35;
            ExcelInteriorColors[36].ColorRGB = Color.FromArgb(255, 255, 153); ExcelInteriorColors[36].Name = "Light Yellow"; ExcelInteriorColors[36].ColorIndex = 36;
            ExcelInteriorColors[37].ColorRGB = Color.FromArgb(153, 204, 255); ExcelInteriorColors[37].Name = "Pale Blue"; ExcelInteriorColors[37].ColorIndex = 37;
            ExcelInteriorColors[38].ColorRGB = Color.FromArgb(255, 153, 204); ExcelInteriorColors[38].Name = "Rose"; ExcelInteriorColors[38].ColorIndex = 38;
            ExcelInteriorColors[39].ColorRGB = Color.FromArgb(204, 153, 255); ExcelInteriorColors[39].Name = "Lavender"; ExcelInteriorColors[39].ColorIndex = 39;
            ExcelInteriorColors[40].ColorRGB = Color.FromArgb(255, 204, 153); ExcelInteriorColors[40].Name = "Tan"; ExcelInteriorColors[40].ColorIndex = 40;
            ExcelInteriorColors[41].ColorRGB = Color.FromArgb(51, 102, 255); ExcelInteriorColors[41].Name = "Light Blue"; ExcelInteriorColors[41].ColorIndex = 41;
            ExcelInteriorColors[42].ColorRGB = Color.FromArgb(51, 204, 204); ExcelInteriorColors[42].Name = "Aqua"; ExcelInteriorColors[42].ColorIndex = 42;
            ExcelInteriorColors[43].ColorRGB = Color.FromArgb(153, 204, 0); ExcelInteriorColors[43].Name = "Lime"; ExcelInteriorColors[43].ColorIndex = 43;
            ExcelInteriorColors[44].ColorRGB = Color.FromArgb(255, 204, 0); ExcelInteriorColors[44].Name = "Gold"; ExcelInteriorColors[44].ColorIndex = 44;
            ExcelInteriorColors[45].ColorRGB = Color.FromArgb(255, 153, 0); ExcelInteriorColors[45].Name = "Light Orange"; ExcelInteriorColors[45].ColorIndex = 45;
            ExcelInteriorColors[46].ColorRGB = Color.FromArgb(255, 102, 0); ExcelInteriorColors[46].Name = "Orange"; ExcelInteriorColors[46].ColorIndex = 46;
            ExcelInteriorColors[47].ColorRGB = Color.FromArgb(102, 102, 153); ExcelInteriorColors[47].Name = "Blue-Gray"; ExcelInteriorColors[47].ColorIndex = 47;
            ExcelInteriorColors[48].ColorRGB = Color.FromArgb(150, 150, 150); ExcelInteriorColors[48].Name = "Gray-40%"; ExcelInteriorColors[48].ColorIndex = 48;
            ExcelInteriorColors[49].ColorRGB = Color.FromArgb(0, 51, 102); ExcelInteriorColors[49].Name = "Dark Teal"; ExcelInteriorColors[49].ColorIndex = 49;
            ExcelInteriorColors[50].ColorRGB = Color.FromArgb(51, 153, 102); ExcelInteriorColors[50].Name = "Sea Green"; ExcelInteriorColors[50].ColorIndex = 50;
            ExcelInteriorColors[51].ColorRGB = Color.FromArgb(0, 51, 0); ExcelInteriorColors[51].Name = "Dark Green"; ExcelInteriorColors[51].ColorIndex = 51;
            ExcelInteriorColors[52].ColorRGB = Color.FromArgb(51, 51, 0); ExcelInteriorColors[52].Name = "Olive Green"; ExcelInteriorColors[52].ColorIndex = 52;
            ExcelInteriorColors[53].ColorRGB = Color.FromArgb(153, 51, 0); ExcelInteriorColors[53].Name = "Brown"; ExcelInteriorColors[53].ColorIndex = 53;
            ExcelInteriorColors[54].ColorRGB = Color.FromArgb(153, 51, 102); ExcelInteriorColors[54].Name = "Plum"; ExcelInteriorColors[54].ColorIndex = 54;
            ExcelInteriorColors[55].ColorRGB = Color.FromArgb(51, 51, 153); ExcelInteriorColors[55].Name = "Indigo"; ExcelInteriorColors[55].ColorIndex = 55;
            ExcelInteriorColors[56].ColorRGB = Color.FromArgb(51, 51, 51); ExcelInteriorColors[56].Name = "Gray-80%"; ExcelInteriorColors[56].ColorIndex = 56;
        }
        public string Name(int Index)
        {
            return ExcelInteriorColors[Index].Name;
        }
        public Color ColorRGB(int Index)
        {
            return ExcelInteriorColors[Index].ColorRGB;
        }
        public int ColorIndex(int Index)
        {
            return ExcelInteriorColors[Index].ColorIndex;
        }
        public int CountColor()
        {
            return 56; // InteriorColors.Length
        }

    }
}
