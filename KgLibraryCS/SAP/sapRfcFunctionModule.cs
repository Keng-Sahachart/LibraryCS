using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using SAP.Middleware.Connector;
using System.Windows.Forms;

using System.Collections;
using System.Windows;
using System.Data;

namespace kgLibraryCs.SAP
{
    /***********************************************/
    public class SapRfcFunctionModule
    {
        public RfcDestination _ecc_SAP;
        private RfcRepository repository; // = _ecc_SAP.Repository
        private IRfcFunction bapiMethod; // For Set Parameter


        private InMemoryDestinationConfiguration objDestConfig = new InMemoryDestinationConfiguration();

        public Dictionary<string, IRfcTable> IRfcTable_Dictionary = new Dictionary<string, IRfcTable>(); // เพื่อความสะดวกในการเรียกชื่อ Table ใช้ Dictionary

        // ----------------------------------------------------------
        // Dim IDestinationConfig As New ECCDestinationConfig
        // Sub New(Optional ByVal ClientName As String = "dev", Optional ByVal ConnectNow As Boolean = True)
        // Try   '" การเชื่อมต่อ ควร สร้างเป็น  Global ไว้ เพื่อ ถ้า สร้าง ซ้ำซ้อน จะ Error

        // RfcDestinationManager.RegisterDestinationConfiguration(IDestinationConfig)
        // _ecc_SAP = RfcDestinationManager.GetDestination(ClientName) '("prd") '("qas") '("dev") '
        // 'tbsaptosql.ecc = _ecc

        // If ConnectNow = True Then
        // repository = _ecc_SAP.Repository
        // End If
        // Catch ex As Exception
        // MessageBox.Show("ไม่สามารถติดต่อ sap database ได้แจ้งผู้ดูแลระบบ :" & ex.Message, "error message", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        // End Try
        // End Sub
        // ----------------------------------------------------------

        public SapRfcFunctionModule(string SystemID = "PRD", string Client = "900", string SystemNumber = "00", string AppServerHostIp = "172.16.22.1"
              , string User = "RFC", string Password = "cvv@2015", string Language = "EN", string SAPRouter = "/H/112.121.138.183/W/cvv/H/")
        {
            try
            {
                RfcConfigParameters parms = new RfcConfigParameters();
                parms[RfcConfigParameters.Name] = SystemID; // "PRD"
                parms[RfcConfigParameters.SystemID] = SystemID; // "PRD"
                parms[RfcConfigParameters.Client] = Client; // "900"
                parms[RfcConfigParameters.SystemNumber] = SystemNumber; // "00"
                parms[RfcConfigParameters.AppServerHost] = AppServerHostIp; 
                     
                parms[RfcConfigParameters.User] = User; 
                parms[RfcConfigParameters.Password] = Password; 
                parms[RfcConfigParameters.Language] = Language; // "EN"
                parms.Add(RfcConfigParameters.SAPRouter, SAPRouter); // 

                RfcDestinationManager.RegisterDestinationConfiguration(objDestConfig);
                objDestConfig.AddOrEditDestination(parms);

                _ecc_SAP = RfcDestinationManager.GetDestination(SystemID); // "PRD"
            }
            // _ecc_SAP.Ping()

            catch (Exception ex)
            {
                MessageBox.Show("Logon failed.  " + ex.Message); 
            }
            finally
            {
            }
        }

        public bool ConnectSAP()
        {
            try
            {
                repository = _ecc_SAP.Repository;
                return true;
            }
            catch //(Exception ex)//Exception ex)
            {
                return false;
            }
        }
        public bool CheckStatusSAP()
        {
            try
            {
                _ecc_SAP.Ping();
                return true;
            }
            catch //(Exception ex)
            {
                return false;
            }
        }

        public void ClearIRfcFunction()
        {
            bapiMethod = null;
            IRfcTable_Dictionary.Clear();
        }

        /// <summary>
        ///     ''' กำหนดชื่อ Function Module ที่จะเรียกใช้งาน
        ///     ''' </summary>
        ///     ''' <param name="FunctionModuleName">ชื่อฟังก์ชั่น Module ใน SAP</param>
        ///     ''' <remarks></remarks>
        public void Step1_CreateFunctionRFC(string FunctionModuleName)
        {
            bapiMethod = repository.CreateFunction(FunctionModuleName);
        }

        /// <summary>
        ///     ''' กำหนด Parameter ที่เป็น  Import ใน Function Module
        ///     ''' </summary>
        ///     ''' <param name="ParameterName">ชื่อ Parameter</param>
        ///     ''' <param name="Input_Parameter">ค่าที่จะใส่เข้าไปใน Parameter</param>
        ///     ''' <remarks></remarks>
        public void Step2_SetImportParameter(string ParameterName, string Input_Parameter)
        {
            bapiMethod.SetValue(ParameterName, Input_Parameter); // 'Import
        }

        /// <summary>
        ///     ''' กำหนด Parameter ที่เป็น Table ที่เป็นทั้ง Import-Export ในตัว ใน Function Module
        ///     ''' / การใส่ข้อมูล ที่เป็น Range ต้องใน เงื้อนไข ให้ครบนะ อย่าลืม  Append ก่อน ใส่ข้อมูล
        ///     ''' / เซ็ต Table ได้ครั้งเดียว
        ///     ''' </summary>
        ///     ''' <param name="TableName">ชื่อ Table</param>
        ///     ''' <returns></returns>
        ///     ''' <remarks></remarks>
        public IRfcTable Step2_SetParameterIRfcTable(string TableName)
        {
            IRfcTable tb_Return = bapiMethod.GetTable(TableName);

            // if ป้องกัน การ Error เนื่องจาก มีการกำหนด Table ไว้แล้ว - ถ้ามีแล้ว ให้ลบออก
            if (IRfcTable_Dictionary.ContainsKey(TableName) == true)
                IRfcTable_Dictionary.Remove(TableName);
            IRfcTable_Dictionary.Add(TableName, tb_Return);

            return (tb_Return);
        }

        /// <summary>
        ///     ''' ทำการ Append Line ข้อมูลก่อน* ใส่ข้อมูล - SetValue  /  **สำคัญคือต้อง Append ก่อน
        ///     ''' </summary>
        ///     ''' <param name="IRFCTableName">ชื่อ Table ที่จะ Append</param>
        ///     ''' <remarks></remarks>
        public void Step3_AppendLineBeforeAdd(string IRFCTableName)
        {
            GetIRfcTable(IRFCTableName).Append();
        }
        /// <summary>
        ///     ''' ใส่ข้อมูลเข้าไปใส่ Line ตาม Parameter Name 
        ///     ''' </summary>
        ///     ''' <param name="IRFCTableName">Table name</param>
        ///     ''' <param name="ValName">ชื่อของ Field ที่จะใส่ข้อมูล</param>
        ///     ''' <param name="Value">ข้อมูลที่จะใส่เข้าไป</param>
        ///     ''' <remarks></remarks>
        public void Step3_SetValueInTable(string IRFCTableName, string ValName, string Value)
        {
            // Dim IRFCTable As IRfcTable
            // IRFCTable = IRfcTable_Dictionary(IRFCTableName)
            // IRFCTable.Append() '  ก่อนที่จะใส่ ค่า ต่อ 1 Row
            GetIRfcTable(IRFCTableName).SetValue(ValName, Value);
        }

        /// <summary>
        ///     ''' ทำการ Retrive หรือ สั่งให้ Function Module ทำงานดึงข้อมูลออกมาใช้งาน
        ///     ''' </summary>
        ///     ''' <remarks></remarks>
        public void Step4_RetriveDataFromRFC()
        {
            bapiMethod.Invoke(_ecc_SAP);
        }

        public DataTable Step5_GetDataTable(string IRFCTableName)
        {
            return IRfcTableExtentions.ToDataTable(IRfcTable_Dictionary[IRFCTableName], IRFCTableName);
        }

        public IRfcTable GetIRfcTable(string IRfcTableName)
        {
            return IRfcTable_Dictionary[IRfcTableName];
        }

        public void Disconnect()
        {
            _ecc_SAP = null;
            RfcDestinationManager.UnregisterDestinationConfiguration(objDestConfig); // (IDestinationConfig)
        }
    }




    /***********************************************/


internal partial class InMemoryDestinationConfiguration : IDestinationConfiguration
{
    private Dictionary<string, RfcConfigParameters> availableDestinations;

    public InMemoryDestinationConfiguration()
    {
        availableDestinations = new Dictionary<string, RfcConfigParameters>();
    }

    public RfcConfigParameters GetParameters(string destinationName)
    {
        RfcConfigParameters foundDestination = null;// default;
        availableDestinations.TryGetValue(destinationName, out foundDestination);
        return foundDestination;
    }

    public bool ChangeEventsSupported()
    {
        return true;
    }

    public event RfcDestinationManager.ConfigurationChangeHandler ConfigurationChanged;

    public void AddOrEditDestination(RfcConfigParameters parameters)
    {
        string name = parameters[RfcConfigParameters.Name];
        if (availableDestinations.ContainsKey(name))
        {
            var EventArgs = new RfcConfigurationEventArgs(RfcConfigParameters.EventType.CHANGED, parameters);
           ;
 /*#error Cannot convert RaiseEventStatementSyntax - see comment for details
             Cannot convert RaiseEventStatementSyntax, System.InvalidCastException: Unable to cast object of type 'Microsoft.CodeAnalysis.CSharp.Syntax.EmptyStatementSyntax' to type 'Microsoft.CodeAnalysis.CSharp.Syntax.ArgumentListSyntax'.
               at ICSharpCode.CodeConverter.CSharp.MethodBodyExecutableStatementVisitor.VisitRaiseEventStatement(RaiseEventStatementSyntax node) in D:\GitWorkspace\CodeConverter\CodeConverter\CSharp\MethodBodyExecutableStatementVisitor.cs:line 404
               at ICSharpCode.CodeConverter.CSharp.HoistedNodeStateVisitor.AddLocalVariables(VisualBasicSyntaxNode node) in D:\GitWorkspace\CodeConverter\CodeConverter\CSharp\HoistedNodeStateVisitor.cs:line 47
               at ICSharpCode.CodeConverter.CSharp.CommentConvertingMethodBodyVisitor.DefaultVisitInnerAsync(SyntaxNode node) in D:\GitWorkspace\CodeConverter\CodeConverter\CSharp\CommentConvertingMethodBodyVisitor.cs:line 29

            Input:
                        RaiseEvent ConfigurationChanged(name, EventArgs)

             */
            if (ConfigurationChanged != null){
                ConfigurationChanged(name, EventArgs);
            }

        }

        availableDestinations[name] = parameters;
        string tmp = "Application server";
        bool isLoadValancing = parameters.TryGetValue(RfcConfigParameters.LogonGroup,out tmp);
        if (isLoadValancing)
        {
            tmp = "Load balancing";
        }
    }

    public void RemoveDestination(string name)
    {
        if (availableDestinations.Remove(name))
        {
            ;
 /*#error Cannot convert RaiseEventStatementSyntax - see comment for details
            Cannot convert RaiseEventStatementSyntax, System.InvalidCastException: Unable to cast object of type 'Microsoft.CodeAnalysis.CSharp.Syntax.EmptyStatementSyntax' to type 'Microsoft.CodeAnalysis.CSharp.Syntax.ArgumentListSyntax'.
               at ICSharpCode.CodeConverter.CSharp.MethodBodyExecutableStatementVisitor.VisitRaiseEventStatement(RaiseEventStatementSyntax node) in D:\GitWorkspace\CodeConverter\CodeConverter\CSharp\MethodBodyExecutableStatementVisitor.cs:line 404
               at ICSharpCode.CodeConverter.CSharp.HoistedNodeStateVisitor.AddLocalVariables(VisualBasicSyntaxNode node) in D:\GitWorkspace\CodeConverter\CodeConverter\CSharp\HoistedNodeStateVisitor.cs:line 47
               at ICSharpCode.CodeConverter.CSharp.CommentConvertingMethodBodyVisitor.DefaultVisitInnerAsync(SyntaxNode node) in D:\GitWorkspace\CodeConverter\CodeConverter\CSharp\CommentConvertingMethodBodyVisitor.cs:line 29

            Input:

                        RaiseEvent ConfigurationChanged(name, New RfcConfigurationEventArgs(RfcConfigParameters.EventType.DELETED))

             */
            if (ConfigurationChanged != null)
            {
                ConfigurationChanged(name, new RfcConfigurationEventArgs(RfcConfigParameters.EventType.DELETED));
            }
        }
    }
}
}


