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
                if (ConfigurationChanged != null)
                {
                    ConfigurationChanged(name, EventArgs);
                }

            }

            availableDestinations[name] = parameters;
            string tmp = "Application server";
            bool isLoadValancing = parameters.TryGetValue(RfcConfigParameters.LogonGroup, out tmp);
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
