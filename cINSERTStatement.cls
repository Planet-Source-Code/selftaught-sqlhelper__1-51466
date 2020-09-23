VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cINSERTStatement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Multiple Record Insert (From SELECT Statement) ONLY

Private mCollSubstituteFields As Collection
Private moSELECTStatement     As cSELECTStatement

Private msTableName           As String
Private msExternalDB          As String

Public Sub AddSubstituteField(psNewName As String, psOldName As String)
    
    'This method causes the 'Old' Field (in the table on which the query
    'was run) to be inserted into the 'New' Field (the field into which
    'records are being inserted)
    
    'If the fields are being inserted into fields with the same names,
    'you do not have to call this method.
    
    mCollSubstituteFields.Add psNewName, psOldName
End Sub

Public Sub Clear()
    Set mCollSubstituteFields = New Collection
    moSELECTStatement.Clear
End Sub

Public Property Let TableName(psVal As String)
    msTableName = psVal
End Property
Public Property Get TableName() As String
    TableName = msTableName
End Property

Public Property Let InsertIntoTable(psVal As String)
    msExternalDB = psVal
End Property
Public Property Get InsertIntoTable() As String
    InsertIntoTable = msExternalDB
End Property

Public Property Get SELECTStatement() As cSELECTStatement
    Set SELECTStatement = moSELECTStatement
End Property

Public Property Get SQLText() As String
    SQLText = sqlStatementInsert(msTableName, _
                                 msExternalDB, _
                                 mCollSubstituteFields, _
                                 moSELECTStatement)
End Property

Private Sub Class_Initialize()
    Set mCollSubstituteFields = New Collection
    Set moSELECTStatement = New cSELECTStatement
End Sub

Private Sub Class_Terminate()
    Set mCollSubstituteFields = Nothing
    Set moSELECTStatement = Nothing
End Sub
