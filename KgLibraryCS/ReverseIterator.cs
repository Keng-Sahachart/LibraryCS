/// <summary>
/// Class ที่เอาไว้ใช้กับ For each  ลูป   ท้าย --> หน้า
/// โดยปกติ  For each ใช้วนตาม Item ไปหน้าอย่างเดียว  หน้า --> ท้าย
/// แต่ Class นี้ ใช้สำหรับ ทำ For Each วนกลับ ท้าย --> หน้า
/// Credit : http://www.devx.com/vb2themax/Tip/18796
/// </summary>
using System.Collections;

namespace kgLibraryCs
{
    public class ReverseIterator : IEnumerable
    {

        // a low-overhead ArrayList to store references
        private ArrayList items = new ArrayList();

        public ReverseIterator(IEnumerable collection)
        {
            // load all the items in the ArrayList, but in reverse order
            foreach (object o in collection)
                items.Insert(0, o);
        }

        public IEnumerator GetEnumerator()
        {
            // return the enumerator of the inner ArrayList
            return items.GetEnumerator();
        }
    }
}

// ######################################################################
// การใช้งาน VB
// ' use an array in this simple test
// Dim arr() As Integer = {1, 2, 3, 4, 5}
// Dim i As Integer
// ' visit array elements in reverse order
// For Each i In New ReverseIterator(arr)
// Console.WriteLine(i)
// Next
// ######################################################################
