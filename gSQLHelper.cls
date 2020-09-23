VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "gSQLHelper"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function sqlFormatValue(pvValue As Variant, _
                Optional ByVal piOperators As eSQLOperators) _
                As String
                
    sqlFormatValue = mSQL.sqlFormatValue(pvValue, piOperators)
End Function

Private Function sqlInOperator(psFieldName As String, pvInObject) _
                 As String
    sqlInOperator = mSQL.sqlInOperator(psFieldName, pvInObject)
End Function

Public Function sqlFieldExpression(psFieldExp As String, _
                          Optional psAlias As String, _
                    Optional ByVal piOperator As eSQLAggregates) _
                As String
    
    sqlFieldExpression = mSQL.sqlFieldExpression(psFieldExp, psAlias, piOperator)

End Function

Public Function sqlFieldCompare(ByVal psField As String, _
                                ByVal pvCompareTo As Variant, _
                                ByVal piOperators As eSQLOperators, _
                       Optional ByVal piParentheses As Long, _
                       Optional ByVal pbFormatted As Boolean) _
                As String
    
    sqlFieldCompare = mSQL.sqlFieldCompare(psField, _
                                           pvCompareTo, _
                                           piOperators, _
                                           piParentheses, _
                                           pbFormatted)
                                           
End Function

Public Function sqlConnectionString(psPath As String, _
                           Optional psUser As String = "Admin", _
                           Optional psPassword As String, _
                           Optional psMode As String = "Share Deny None", _
                           Optional psJetDBPassword As String, _
                     Optional ByVal pbJetEncrypt As Boolean, _
                           Optional psProvider As String = "Microsoft.Jet.OLEDB.4.0") _
                As String
    sqlConnectionString = "Provider=" & psProvider & ";Password=""" & psPassword & """;User ID=" & psUser & ";Data Source=""" & psPath & """;Mode=" & psMode
    If Len(psJetDBPassword) > 0 Then sqlConnectionString = sqlConnectionString & ";Jet OLEDB:Database Password=""" & psJetDBPassword & """"
    If pbJetEncrypt Then sqlConnectionString = sqlConnectionString & ";Jet OLEDB:Encrypt Database=" & pbJetEncrypt
End Function

