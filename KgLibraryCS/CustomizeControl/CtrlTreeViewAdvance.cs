// สร้างจาก
// http://support.microsoft.com/kb/311318/th
// imports System.Windows.Forms.TreeView
using System.Linq;
using System.Collections.Generic;
using System.Collections;
using System;
using System.Windows.Forms;
using System.Text.RegularExpressions;
using System.Drawing;
using Microsoft.VisualBasic.CompilerServices;

namespace KengsLibraryCs
{

    // Customize เพิ่มเติม
    // เพิ่มคุณสมบัติ
    // -ค้นหาข้อมูลของ Node ได้ จาก text,name,key,value ส่งออกมาเป็น list of node
    // -เคลียร์สี พื้นหลังของ Node จากการ find แล้วเอาไปใส่สีเอง
    // -ทำให้เก็บข้อมูลอื่นๆในรูปแบบ Tag (String) เช่น [Var1=Val1][Var2=Val2] ได้ ทำให้เก็บข้อมูลได้เป็นตัวแปรไม่จำกัด โดยการใช้ regex ช่วย เพิ่มเติมหรือแก้ไข ตัวแปร tag นั้นๆ 
    // -ค้นหาข้อมูลของ Node ได้จาก Tag
    public partial class CtrlTreeViewAdvance : TreeView
    {
        // Inherits TreeView

        // Public Overridable Overloads Property SelectedNode As TreeViewAdvance.TreeNode
        // 'Implements TreeViewAdvance.SelectedNode

        // Get
        // Return Me.SelectedNode
        // End Get

        // Set(ByVal Node As TreeNodeCollection)
        // Me.SelectedNode = TreeNode
        // End Set
        // End Property

        // ###################################################################################################
        // ค่าให้เลือกสำหรับกำหนดการค้นหาข้อมูล
        enum ChoiceFindIn
        {
            InName = 0,
            InText = 1,
            InNodeKey = 2,
            InNodeValue = 3,
            InNodeValueTag = 4
        }
        /// <summary> ค้นหา แบบ ฟังก์ชั่น nodes.Find ได้แล้ว ดัดแปลงมาจาก CallRecursive  ได้กลับมาเป็น Array </summary>
    /// <param name="aTreeView">TreeView ที่จะค้นหา</param>
    /// <param name="TextFind">คำที่ต้องการค้นหา</param>
    /// <param name="SeachAllChildNode">Logic ว่าจะค้นหาใน Child Node ด้วยหรือป่าว</param>
    /// <param name="ChoiceFindIn">เลือกว่าจะค้นหาอะไร ว่าง(0):Name, 1:Text, 2:NodeKey, 3:NodeValue</param>
    /// <returns>ส่งออกเป็น Array เลือกแก้ไข ได้ว่าจะเอาเป็น List ก็ได้</returns>
        public object FindNode(TreeView aTreeView, string TextFind
                                 , bool SeachAllChildNode = true
                                , ChoiceFindIn ChoiceFindIn = ChoiceFindIn.InName)
        {
            var NodesDetected = new List<TreeNode>(); // เก็บโหนดที่พบ เพื่อ Return

            foreach (TreeNode Nde in aTreeView.Nodes)
            {
                string ValForCheck; // เลือกข้อมูลที่จะค้นหา
                switch (ChoiceFindIn)
                {
                    case (ChoiceFindIn)1:
                        {
                            ValForCheck = Nde.Text;
                            break;
                        }

                    case (ChoiceFindIn)2:
                        {
                            ValForCheck = Nde.NodeKey;
                            break;
                        }

                    case (ChoiceFindIn)3:
                        {
                            ValForCheck = Nde.NodeValue;
                            break;
                        }

                    default:
                        {
                            ValForCheck = Nde.Name;
                            break;
                        }
                }

                if (SeachAllChildNode == true)
                    FindChildNode(ref Nde, ref NodesDetected, TextFind, ChoiceFindIn); // ฟังก์ชั่นสำหรับ ค้นหาใน โหนดลูก
                else if (ValForCheck.Contains(TextFind))
                    NodesDetected.Add(Nde);// เก็บโหนดที่พบ
            }
            // ตอนส่งออกไป ต้องเอาตัวแปร Array หรือ List มารับ
            // Return NodesDetected ' กรณีนี้ ต้องเอา ตัวแปร List(Of TreeNode) มารับ
            var TrNdeDetect = NodesDetected.ToArray(); // กรณีนี้ สำหรับ เอา array ของ TreeNode มารับ
            return TrNdeDetect;
        }
        /// <summary> ฟังก์ชั่นลูกของ FindNode ค้นหาคำ ใน Node ChildNode </summary>
    /// <param name="Nde">(ByRef) Node ที่มี ChildNode</param>
    /// <param name="NodesDetected">(ByRef) ตัวแปร List(Of TreeNode) ที่ส่งเข้ามาเพื่อรับ Node ที่ "ค้นเจอ" ออกไป</param>
    /// <param name="TextFind">คำที่ต้องการค้นหา</param>
        private void FindChildNode(ref TreeNode Nde, ref List<TreeNode> NodesDetected
                                 , string TextFind, ChoiceFindIn ChoiceFindIn = ChoiceFindIn.InName)
        {
            string ValForCheck; // เลือกข้อมูลที่จะค้นหา
            switch (ChoiceFindIn)
            {
                case (ChoiceFindIn)1:
                    {
                        ValForCheck = Nde.Text;
                        break;
                    }

                case (ChoiceFindIn)2:
                    {
                        ValForCheck = Nde.NodeKey;
                        break;
                    }

                case (ChoiceFindIn)3:
                    {
                        ValForCheck = Nde.NodeValue;
                        break;
                    }

                default:
                    {
                        ValForCheck = Nde.Name;
                        break;
                    }
            }
            if (ValForCheck.Contains(TextFind))
                NodesDetected.Add(Nde);
            foreach (TreeNode aNode in Nde.Nodes)
                FindChildNode(ref aNode, ref NodesDetected, TextFind, ChoiceFindIn);
        }
        // ######################
        // OverLoad Function
        // ######################
        /// <summary> ค้นหาจาก ตัวแปรที่ กำหนดให้เก็บ แบบ Tag</summary>
    /// <param name="TagName">ชื่อของ Tag ที่ถูกกำหนดไว้</param>
    /// <param name="TextFind">คำที่ต้องการค้นหา</param>
    /// <param name="SeachAllChildNode">Logic ว่าจะค้นหาใน Child Node ด้วยหรือป่าว</param>
    /// <returns>ส่งออกเป็น Array เลือกแก้ไข ได้ว่าจะเอาเป็น List ก็ได้</returns>
        public object FindNode(string TagName, string TextFind
                                 , bool SeachAllChildNode = true)
        {
            var NodesDetected = new List<TreeNode>(); // เก็บโหนดที่พบ เพื่อ Return

            foreach (TreeNode Nde in Nodes)
            {
                // จะค้นหา Node ลูกด้วยหรือไม่
                if (SeachAllChildNode == true)
                    FindChildNode(ref Nde, ref NodesDetected, TagName, TextFind); // ฟังก์ชั่นสำหรับ ค้นหาใน โหนดลูก
                else if (Nde.get_NodeValueTag(TagName).Contains(TextFind) != null)
                    NodesDetected.Add(Nde);// เก็บโหนดที่พบ
            }
            // ตอนส่งออกไป ต้องเอาตัวแปร Array หรือ List มารับ
            // Return NodesDetected ' กรณีนี้ ต้องเอา ตัวแปร List(Of TreeNode) มารับ
            var TrNdeDetect = NodesDetected.ToArray(); // กรณีนี้ สำหรับ เอา array ของ TreeNode มารับ
            return TrNdeDetect;
        }
        /// <summary> ฟังก์ชั่นลูกของ FindNode ค้นหาคำ ใน ตัวแปรแบบ Tag ใน Node ChildNode </summary>
    /// <param name="Nde">(ByRef) Node ที่มี ChildNode</param>
    /// <param name="NodesDetected">(ByRef) ตัวแปร List(Of TreeNode) ที่ส่งเข้ามาเพื่อรับ Node ที่ "ค้นเจอ" ออกไป</param>
    /// <param name="TextFind">คำที่ต้องการค้นหา</param>
        private void FindChildNode(ref TreeNode Nde, ref List<TreeNode> NodesDetected
                                 , string TagName, string TextFind)
        {
            if (Operators.ConditionalCompareObjectEqual(Nde.get_NodeValueTag(TagName).Contains(TextFind), true, false))
                NodesDetected.Add(Nde);// เก็บโหนดที่พบ

            foreach (TreeNode aNode in Nde.Nodes)
                FindChildNode(ref aNode, ref NodesDetected, TagName, TextFind);
        }
        // ###################################################################################################
        // อาจไม่จำเป็นต้องใช้แล้ว เพราะ มีฟังก์ชั่น เรียกใช้ทุกโหนด
        public void ClearNodeBackColor(bool ClearAllChildNode = true)
        {
            foreach (TreeNode Nde in Nodes)
            {
                if (ClearAllChildNode == true)
                    ClearChildNodeColor(ref Nde);
                else
                    Nde.BackColor = Color.Empty;
            }
        }
        /// <summary> ฟังก์ชั่นลูกของ ClearNodeBackColor เพื่อเคลียร์สีของ TreeNode ลูก </summary>
        private void ClearChildNodeColor(ref TreeNode Nde)
        {
            Nde.BackColor = Color.Empty;
            foreach (TreeNode aNode in Nde.Nodes) // ถ้ามีโหนดย่อยก็จะเข้าไปใน Loop เอง
                ClearChildNodeColor(ref aNode);
        }
        // ###################################################################################################
        /// <summary>สำหรับ นำทุกโหนดไปใช้ทำอะไรซักอย่าง เช่น เคลียร์สี หรืออื่นๆ</summary>
        public object GetAllNodeInArray()
        {
            var NodesDetected = new List<TreeNode>(); // เก็บผลลัพพ์
            foreach (TreeNode Nde in Nodes)
            {
                NodesDetected.Add(Nde);
                GetAllChildNodeInArray(ref Nde, ref NodesDetected);
            }
            return NodesDetected.ToArray();
        }
        /// <summary>
    /// ฟังก์ชั่นลูกของ GetAllNodeInArray เพื่อเอา Node ลูกออกมา
    /// </summary>
    /// <param name="Nde">TreeNode</param>
    /// <param name="ReturnNode">ตัวแปร List(Of TreeViewAdvance.TreeNode) แบบ ByRef เอาเข้าไปแล้วเอาค่าออกมา</param>
        private void GetAllChildNodeInArray(ref TreeNode Nde, ref List<TreeNode> ReturnNode)
        {
            foreach (TreeNode aNode in Nde.Nodes)
            {
                ReturnNode.Add(aNode);
                GetAllChildNodeInArray(ref aNode, ref ReturnNode);
            }
        }

        /// <summary>
    /// ใช้ดึง Node ทั้งหมดที่อยู่ใน Path ของ Node ที่จะดู ตั้งแต่ Node ที่ใส่มา ขึ้นไป
    /// </summary>
    /// <param name="fromNode">Node ที่ต้องการหา Parent ทั้งหมด</param>
    /// <returns>จะส่ง Node ออกมาตามลำดับ Revers ในสุดไปนอกสุด</returns>
    /// <remarks></remarks>
        public IEnumerable<TreeNode> GetAllParentNodes(TreeNode fromNode)
        {
            var result = new List<TreeNode>();
            while (fromNode != null)
            {
                result.Add(fromNode);
                fromNode = fromNode.Parent;
            }
            return result;
        }

        /// <summary>
    /// ยังไม่สรุปว่าจะใช้ทำอะไร
    /// </summary>
        public void GetAllNodeInArray2(ref TreeNode Node, ref List<TreeNode> InputToReturnNode)
        {
            foreach (TreeNode Nde in Nodes)
            {
                InputToReturnNode.Add(Nde);
                GetAllNodeInArray2(ref Nde, ref InputToReturnNode);
            }
        }
        // ###################################################################################################
        /// Event สำหรับ เหตุการ Lost Focus.
    /// เพราะ Control ปกติ เวลา LostFocus มันไม่ HighLight ที่ Node ที่ได้เลือกไว้เลย อาจทำให้สับสน ว่าเลือก Item ไหนอยู่
        private void TreeView1_LostFocus(object sender, EventArgs e)
        {
            if (SelectedNode != null)
            {
                if (SelectedNode.BackColor == Color.Yellow)
                    SelectedNode.BackColor = Color.CadetBlue;
                else
                    SelectedNode.BackColor = SystemColors.InactiveCaption;
            }
        }
        private void TreeView1_GotFocus(object sender, EventArgs e)
        {
            if (SelectedNode != null)
            {
                // When the TreeView receives focus again the last selected nodes backcolor needs to be reset.
                if (SelectedNode.BackColor != Color.Yellow)
                {
                    if (SelectedNode.BackColor == Color.CadetBlue)
                        SelectedNode.BackColor = Color.Yellow;
                    else
                        SelectedNode.BackColor = Color.Empty;
                }
            }
        }
        // ###################################################################################################
        // ###################################################################################################
        // ###################################################################################################
        public class TreeNode : System.Windows.Forms.TreeNode, IDictionaryEnumerator
        {
            private DictionaryEntry nodeEntry;
            private IEnumerator enumerator;

            public object ValStrctr; // สำหรับเก็บตัวแปร Structure เผื่อกรณี ต้องการตัวแปรเพิ่ม

            // สำหรับเก็บ ข้อมูล ในรูปแบบ Tag [ValName1=Val1][ValName1=Val1] ใช้ regex ในการกระทำกับข้อมูล เก็บข้อมูลไม่จำกัด
            private string ValueTag = ""; // ต้องใส่ "" ไว้ ถ้าเป็น Nothing  regex จะ error

            public TreeNode(string NodeName = null)
            {
                enumerator = Nodes.GetEnumerator();
                nodeEntry.Key = "";
                nodeEntry.Value = "";
                if (NodeName != null)
                    Text = NodeName;
            }

            /// <summary>Key สำหรับเก็บข้อความ เพื่อใช้หรือเอาไว้ค้นหาได้</summary>
            public string NodeKey
            {
                get
                {
                    return nodeEntry.Key.ToString();
                }
                set
                {
                    nodeEntry.Key = value;
                }
            }
            /// <summary>Value สำหรับเก็บข้อความ เพื่อใช้หรือเอาไว้ค้นหาได้</summary>
            public object NodeValue
            {
                get
                {
                    return nodeEntry.Value;
                }
                set
                {
                    nodeEntry.Value = value;
                }
            }
            public string GetAllValTag()
            {
                return ValueTag;
            }
            public object get_NodeValueTag(string TagName)
            {
                // Return nodeEntry.Value
                string Pattern = string.Format(@"\[{0}=(?<DATA>.*?)\]", TagName); // "\[DepType=(?<DATA>.*?)\]"
                var Matches = Regex.Matches(ValueTag, Pattern, RegexOptions.IgnoreCase | RegexOptions.Singleline);
                if (Matches != null && Matches.Count > 0)
                {
                    // For Each Match As Match In Matches 'ตัวอย่างว่าพบหลายตัว
                    if (Matches[0].Groups["DATA"].Success)
                        // Return Match.Groups("DATA").Value
                        return Matches[0].Groups["DATA"].Value;
                }
                return null; // ถ้าไม่มี
            }

            public void set_NodeValueTag(string TagName, object value)
            {
                // nodeEntry.Value = Value
                if (Operators.ConditionalCompareObjectEqual(get_NodeValueTag(TagName), null, false))
                    // ถ้าไม่มีให้เพิ่ม Tag เข้าไป
                    ValueTag += string.Format("[{0}={1}]", TagName, value);
                else
                {
                    // แก้ไข ข้อมูลที่มีอยู่แล้ว
                    string PaternTag = string.Format(@"\[{0}=(.*?)\]", TagName);
                    string AllText = ValueTag;
                    string NewValue = string.Format("[{0}={1}]", TagName, value);
                    ValueTag = Regex.Replace(AllText, PaternTag, NewValue);
                }
            }

            public virtual new DictionaryEntry Entry
            {
                get
                {
                    return nodeEntry;
                }
            }

            public virtual new bool MoveNext()
            {
                bool Success;
                Success = enumerator.MoveNext();
                return Success;
            }

            public virtual new object Current
            {
                get
                {
                    return enumerator.Current;
                }
            }

            public virtual new object Key
            {
                get
                {
                    return nodeEntry.Key;
                }
            }

            public virtual new object Value
            {
                get
                {
                    return nodeEntry.Value;
                }
            }
            // Public Overridable Overloads ReadOnly Property Value() As Object _
            // Implements IDictionaryEnumerator.Value
            // Get
            // Return nodeEntry.Value
            // End Get
            // End Property

            public virtual new void Reset()
            {
                enumerator.Reset();
            }
        }
    }
}
