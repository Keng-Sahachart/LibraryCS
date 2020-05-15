using System.Collections.Generic;
using System.Windows.Forms;

namespace kgLibraryCs
{
    public static class FncTreeView
    {
        /// <summary>
        /// ใช้ดึง Node ทั้งหมดที่อยู่ใน Path ของ Node ที่จะดู --> ตั้งแต่ Node ที่ใส่มา ไต่ Lv ปัจจุบัน ขึ้นไปหา Lv บนสุด
        /// </summary>
        /// <param name="fromNode">Node ที่ต้องการหา Parent ทั้งหมด</param>
        /// <returns>จะส่ง Node ออกมาตามลำดับ Revers ในสุดไปนอกสุด</returns>
        /// <remarks></remarks>
        public static IEnumerable<TreeNode> GetAllParentNodes(TreeNode fromNode)
        {
            var result = new List<TreeNode>();
            while (fromNode != null)
            {
                result.Add(fromNode);
                fromNode = fromNode.Parent;
            }
            return result;
        }
    }
}
