VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cUPDATEStatement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mCollSet As Collection
Private mCollWhere As Collection

Private msTableName As String

Public Sub AddSetClause(psField As String, _
                        pvValue As Variant, _
         Optional ByVal pbFormatted As Boolean = False)
    
    Dim loClause As cClause: Set loClause = New cClause
    
    With loClause
        .Value = pvValue
        .Text = psField
        .Formatted = pbFormatted
    End With
    
    mCollSet.Add loClause

End Sub

Public Sub AddWhereClause(psField As String, _
                          pvValue As Variant, _
           Optional ByVal piOperators As eSQLOperators = sqlEqual, _
           Optional ByVal piParentheses As Long, _
           Optional ByVal pbFormatted As Boolean = False)
           
    Dim loClause As cClause: Set loClause = New cClause
    
    Select Case piOperators Mod sqlOR
        Case Is < sqlEqual
            Err.Raise 5
        Case Is > sqlIn
            Err.Raise 5
    End Select
    
    If mCollWhere.Count > 0 Then
        If Not BitIsSet(piOperators, sqlOR) Or BitIsSet(piOperators, sqlAND) Then _
                SetBit piOperators, sqlAND, True
    Else
        piOperators = piOperators Mod sqlOR
    End If
    
    With loClause
        .Text = psField
        .Value = pvValue
        .Operator = piOperators
        .Parentheses = piParentheses
        .Formatted = pbFormatted
    End With
    
    mCollWhere.Add loClause
    
End Sub

Public Property Get TableName() As String
    TableName = msTableName
End Property
Public Property Let TableName(psVal As String)
    msTableName = psVal
End Property

Public Property Get SQLText() As String
Attribute SQLText.VB_UserMemId = 0
    Dim loEach  As cClause
    Dim i       As Long
    Dim liCount As Long: liCount = mCollSet.Count
    Dim loSet   As cStringBuilder: Set loSet = New cStringBuilder
    Dim loWhere As cStringBuilder: Set loWhere = New cStringBuilder
    
    If liCount > 0 Then
        For Each loEach In mCollSet
            With loEach
                i = i + 1
                loSet.Append sqlFieldCompare(.Text, _
                                             .Value, sqlEqual, _
                                             , _
                                             .Formatted)
                If i < liCount Then loSet.Append ", "
            End With
        Next
    Else
        'Cannot have update statement with nothing to set!
        Debug.Assert False
    End If
    
    If mCollWhere.Count > 0 Then
        For Each loEach In mCollWhere
            With loEach
                loWhere.Append sqlFieldCompare(.Text, _
                                              .Value, _
                                              .Operator, _
                                              .Parentheses, _
                                              .Formatted)
            End With
        Next
    End If
    
    SQLText = sqlStatementUpdate(msTableName, _
                              loSet.ToString, _
                              loWhere.ToString)
   
End Property

Public Sub Clear()
    Set mCollSet = New Collection
    Set mCollWhere = New Collection
End Sub

Private Sub Class_Initialize()
    Set mCollSet = New Collection
    Set mCollWhere = New Collection
End Sub

Private Sub Class_Terminate()
    Set mCollSet = Nothing
    Set mCollWhere = Nothing
End Sub

