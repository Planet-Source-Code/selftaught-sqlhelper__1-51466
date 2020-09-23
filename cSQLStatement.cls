VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cSELECTStatement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum eSQLOperators
    'You may pick one of these:
    sqlEqual = 1
    sqlGreaterThan
    sqlGreaterThanEqualto
    sqlLessThan
    sqlLessThanEqualTo
    sqlNotEqual
    sqlBetween
    sqlLike
    sqlIn
    
    'And one of these:
    sqlOR = 64
    sqlAND = 128
End Enum

Public Enum eSQLAggregates
    sqlAvg = 1
    sqlCount
    sqlMax
    sqlMin
    sqlSum
End Enum

Public Enum eSQLSortModes
    sqlOff
    sqlAscending
    sqlDescending
End Enum

Public Enum eSQLSelectTypes
    sqlAll = 0
    sqlDistinct = -101
    sqlDistinctRow = -102
    'A value from -100 to -1 is interpreted as TOP (abs(x)) PERCENT
    'A value >= 1 is interpreted as TOP x
End Enum

Private mCollSelect   As Collection
Private mCollWhere    As Collection
Private mCollChildren As Collection
Private mCollHaving   As Collection

Private msTableName   As String
Private msSortColumn  As String
Private msGroupBy     As String
Private msSelectInto  As String
Private msExternalDB  As String


Private miSortMode    As eSQLSortModes
Private miSelectType  As Long

'Friendly properties
Private msRelatedOn As String
Private msRelatedTo As String
Private msFieldName As String

Public Sub AddSelectClause(psFieldExp As String, _
            Optional ByVal psAlias As String, _
            Optional ByVal piOperator As eSQLAggregates)
    
    'Call this method for each field that you want to return from the database.
    'FieldExp can be a simple field name or a string containing a SQL calculation.
    
    Dim loClause As cClause: Set loClause = New cClause
    
    With loClause
        .Value = psFieldExp
        .Operator = piOperator
        .Text = psAlias
    End With
    
    mCollSelect.Add loClause
    
End Sub

Public Sub AddWhereClause(psField As String, _
                          pvValue As Variant, _
           Optional ByVal piOperators As eSQLOperators = sqlEqual, _
           Optional ByVal piParentheses As Long, _
           Optional ByVal pbFormatted As Boolean = False)
           
    'Call This method to limit the fields returned from the database.
    
    'You can OR together any one of the other eSQLOperators with sqlOR or sqlAND
    'to indicate the logic used to evaluate the expression.  If this is the
    'first where clause, the and/or part is ignored.
    
    'If piParentheses < 0 then
    '    abs(piParenthases) "("'s are added to the beginning of the expression
    'elseif piparentheses > 0 then
    '    (piParenthases) ")"'s are added to the end.
    'endif
    
    'If pbFormatted = True, then it is assumed that all formatting has been
    'done.  This allows you to create a clause where "psField = FieldName * 2"
    'would not get formatted again as "psField = 'FieldName * 2'"
    
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

Public Sub AddHavingClause(psField As String, _
                           pvValue As Variant, _
            Optional ByVal piOperators As eSQLOperators = sqlEqual, _
            Optional ByVal piParentheses As Long, _
            Optional ByVal pbFormatted As Boolean = False)
    
    'call this method to add having clauses to the Group By Clause.
    
    'Arguments are the same as AddSelectClause
    
    'These clauses do not affect the SQLText property unless the Group By
    'clause is <> ""
    
    Dim loClause As cClause: Set loClause = New cClause
    
    Select Case piOperators Mod sqlOR
        Case Is < sqlEqual
            Err.Raise 5
        Case Is > sqlIn
            Err.Raise 5
    End Select
    
    If mCollHaving.Count > 0 Then
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
    
    mCollHaving.Add loClause
    
End Sub

Public Function AddChild(psChildTable As String, _
                         psRelatedOn As String, _
                         psRelatedTo As String, _
                         psAppendFieldName As String) _
                As cSELECTStatement
    
    'Call this method to append child queries.
    'psChildTable is the table on which to run the query.
    'psRelatedOn is the field within the child query that will be matched to
    'psRelatedTo in the parent query.
    'psAppendFieldName is the alias for the child recordset.
    
    Set AddChild = New cSELECTStatement
    
    With AddChild
        .TableName = psChildTable
        .RelatedOnField = psRelatedOn
        .RelatedToField = psRelatedTo
        .RelatedAppendFieldName = psAppendFieldName
    End With
    
    mCollChildren.Add AddChild, psChildTable

End Function

Public Property Get GroupBy() As String
    GroupBy = msGroupBy
End Property
Public Property Let GroupBy(psVal As String)
    msGroupBy = psVal
End Property

Public Property Get TableName() As String
    TableName = msTableName
End Property
Public Property Let TableName(psVal As String)
    
    'If Len(msRelatedTo) = 0 And Len(msRelatedOn) = 0 Then
        msTableName = psVal '_
    'Else _
        ': Err.Raise 425, "cSELECTStatement.LetTableName", "Invalid Object Use."
        
End Property

Public Property Let SortColumn(psVal As String)
    msSortColumn = psVal
End Property
Public Property Get SortColumn() As String
    msSortColumn = SortColumn
End Property

Public Property Get SortMode() As eSQLSortModes
    SortMode = miSortMode
End Property
Public Property Let SortMode(piVal As eSQLSortModes)
    If piVal < sqlOff Or piVal > sqlDescending Then Err.Raise 5
    miSortMode = piVal
End Property

Public Property Let SelectType(piVal As eSQLSelectTypes)
    miSelectType = piVal
    
'    Public Enum eSQLSelectTypes
'        sqlAll = -101
'        sqlDistinct = -102
'        sqlDistinctRow = -103
'        'A value from -100 to -1 is interpreted as TOP (abs(x)) PERCENT
'        'A value >= 1 is interpreted as TOP x
'    End Enum
    
End Property
Public Property Get SelectType() As eSQLSelectTypes
    SelectType = miSelectType
End Property

Public Property Let SelectIntoTable(psVal As String)
    msSelectInto = psVal
End Property
Public Property Get SelectIntoTable() As String
    SelectIntoTable = msSelectInto
End Property

Public Property Let SelectIntoExternalDB(psVal As String)
    msExternalDB = psVal
End Property
Public Property Get SelectIntoExternalDB() As String
    SelectIntoExternalDB = msExternalDB
End Property

Public Property Get SQLText() As String
Attribute SQLText.VB_UserMemId = 0
    Dim loEach  As cClause
    Dim loWhere As cStringBuilder: Set loWhere = New cStringBuilder
    Dim lsFrom As String: lsFrom = sqlSelect(miSelectType, mCollSelect) & _
                                   sqlSelectInto(msSelectInto, msExternalDB)

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
    
    SQLText = sqlStatementSelect(msTableName, lsFrom, loWhere.ToString) & _
              sqlOrderBy(msSortColumn, miSortMode)
    
    If mCollChildren.Count > 0 Then
        SQLText = sqlShape(SQLText, mCollChildren)
    End If
    
    If Not Len(msGroupBy) = 0 Then
        SQLText = SQLText & sqlGroupBy(msGroupBy, mCollHaving)
    End If
End Property

Public Sub Clear(Optional ClearSelect As Boolean = True, Optional ClearWhere As Boolean = True, Optional ClearGroupByHaving As Boolean = True, Optional ClearChildren As Boolean = True, Optional ClearSort As Boolean = True)
    If ClearSelect Then
        Set mCollSelect = New Collection
        miSelectType = 0
        msExternalDB = ""
        msSelectInto = ""
    End If
    
    If ClearGroupByHaving Then
        Set mCollHaving = New Collection
        msGroupBy = ""
    End If
    
    If ClearSort Then
        miSortMode = 0
        msSortColumn = ""
    End If
    
    If ClearChildren Then Set mCollChildren = New Collection
    If ClearWhere Then Set mCollWhere = New Collection
   
End Sub

Private Sub Class_Initialize()
    Clear
End Sub

Private Sub Class_Terminate()
    Set mCollSelect = Nothing
    Set mCollWhere = Nothing
    Set mCollChildren = Nothing
    Set mCollHaving = Nothing
End Sub

Friend Property Let RelatedOnField(psVal As String)
    msRelatedOn = psVal
End Property
Friend Property Get RelatedOnField() As String
    RelatedOnField = msRelatedOn
End Property

Friend Property Let RelatedToField(psVal As String)
    msRelatedTo = psVal
End Property
Friend Property Get RelatedToField() As String
    RelatedToField = msRelatedTo
End Property

Friend Property Let RelatedAppendFieldName(psVal As String)
    msFieldName = psVal
End Property
Friend Property Get RelatedAppendFieldName() As String
    RelatedAppendFieldName = msFieldName
End Property

Friend Property Get SelectClauses() As Collection
    Set SelectClauses = mCollSelect
End Property
