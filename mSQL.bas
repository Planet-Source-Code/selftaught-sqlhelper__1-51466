Attribute VB_Name = "mSQL"
Option Explicit

Public Enum eStatements
    smtSelect
    smtUpdate
    smtInsert
    smtDelete
End Enum

Private Const sShape = "SHAPE "
Private Const sAppend = " APPEND "
Private Const sRelate = " RELATE "
Private Const sTo = " TO "
Private Const sAs = " AS "

Private Const sUpdate = "UPDATE "
Private Const sSet = " SET "

Private Const sDelete = "DELETE "

Private Const sGroupBy = " GROUP BY "
Private Const sHaving = " HAVING "

Private Const sInsertInto = "INSERT INTO "

Private Const sSelect = "SELECT "
Private Const sInto = " INTO "
Private Const sFrom = " FROM "
Private Const sWhere = " WHERE "
Private Const sOrderBy = " ORDER BY "
Private Const sBetween = " BETWEEN "
Private Const sLike = " LIKE "
Private Const sIn = " IN "
Private Const sAnd = " AND "
Private Const sNot = "NOT"
Private Const sOr = " OR "
Private Const sGT = " > "
Private Const sLT = " < "
Private Const sGTE = " >= "
Private Const sLTE = " <= "
Private Const sEQL = " = "
Private Const sNotEQL = " <> "

Private Const sAvg = "AVG"
Private Const sCount = "COUNT"
Private Const sDistinct = "DISTINCT "
Private Const sDistinctRow = "DISTINCTROW "
Private Const sMax = "MAX"
Private Const sMin = "MIN"
Private Const sSum = "SUM"
Private Const sTopX = "TOP "
Private Const sPercent = " PERCENT "

Public Function sqlSelect(ByVal piSelectType As eSQLSelectTypes, pCollSelect As Collection) _
                As String
    
    Dim i        As Long
    Dim liCount  As Long
    Dim loEach   As cClause
    Dim loReturn As cStringBuilder: Set loReturn = New cStringBuilder
    
    Select Case piSelectType
        Case sqlDistinct
            loReturn.Append sDistinct
        Case sqlDistinctRow
            loReturn.Append sDistinctRow
        Case Is <> sqlAll
            If piSelectType > 0 Then
                loReturn.Append sTopX & piSelectType & " "
            ElseIf piSelectType < 0 Then
                If piSelectType < -100 Then piSelectType = -100
                loReturn.Append sTopX & Abs(piSelectType) & sPercent
            End If
    End Select
    
    i = 1
    liCount = pCollSelect.Count
    If liCount > 0 Then
        For Each loEach In pCollSelect
            With loEach
                If i = liCount Then
                    loReturn.Append sqlFieldExpression(.Value, .Text, .Operator)
                Else
                    loReturn.Append sqlFieldExpression(.Value, .Text, .Operator) & ", "
                End If
                i = i + 1
            End With
        Next
    Else
        loReturn.Append "*"
    End If
    sqlSelect = loReturn.ToString
End Function

Public Function sqlSelectInto(psIntoTable As String, psIntoExternalDB As String) _
                As String
    
    If Len(psIntoTable) > 0 Then
        If Len(psIntoExternalDB) > 0 Then
            sqlSelectInto = sInto & psIntoTable & sIn & "'" & psIntoExternalDB & "'"
        Else
            sqlSelectInto = sInto & psIntoTable
        End If
    End If

End Function

Public Function sqlFieldExpression(psFieldExp As String, _
                          Optional psAlias As String, _
                    Optional ByVal piOperator As eSQLAggregates) _
                As String
    
    Select Case piOperator
        Case 0
            sqlFieldExpression = psFieldExp
        Case sqlAvg
            sqlFieldExpression = sAvg & "(" & psFieldExp & ")"
        Case sqlCount
            sqlFieldExpression = sCount & "(" & psFieldExp & ")"
        Case sqlMax
            sqlFieldExpression = sMax & "(" & psFieldExp & ")"
        Case sqlMin
            sqlFieldExpression = sMin & "(" & psFieldExp & ")"
        Case sqlSum
            sqlFieldExpression = sSum & "(" & psFieldExp & ")"
    End Select
    If Len(psAlias) > 0 Then sqlFieldExpression = sqlFieldExpression & sAs & psAlias
End Function

Public Function sqlFieldCompare(ByVal psField As String, _
                                ByVal pvCompareTo As Variant, _
                                ByVal piOperators As eSQLOperators, _
                       Optional ByVal piParentheses As Long, _
                       Optional ByVal pbFormatted As Boolean) _
                As String

    If BitIsSet(piOperators, sqlAND) Then
        sqlFieldCompare = sAnd
    ElseIf BitIsSet(piOperators, sqlOR) Then
        sqlFieldCompare = sOr
    End If
    
    If piParentheses < 0 Then
        piParentheses = Abs(piParentheses)
        sqlFieldCompare = sqlFieldCompare & String(piParentheses, "(")
        piParentheses = 0
    End If

    Select Case piOperators Mod sqlOR
        Case sqlIn
            sqlFieldCompare = sqlFieldCompare & sqlInOperator(psField, pvCompareTo)
        Case Else
            sqlFieldCompare = sqlFieldCompare & psField & sqlFormatValue(pvCompareTo, piOperators, pbFormatted)
    End Select

    If piParentheses > 0 Then _
        sqlFieldCompare = sqlFieldCompare & String(piParentheses, ")")

End Function

Public Function sqlFormatValue(pvValue As Variant, _
                Optional ByVal piOperators As eSQLOperators, _
                Optional ByVal pbSkipValueFormatting As Boolean) _
                As String
    
    Dim liType      As VbVarType
    Dim lvFormatted As Variant
    
    liType = VarType(pvValue)
    
    If Not (IsNull(pvValue) Or IsEmpty(pvValue)) Then _
        sqlFormatValue = sqlGetOperator(piOperators)
    
    Select Case piOperators Mod sqlOR
        Case sqlBetween
            If BitIsSet(liType, vbArray) Then
                
                sqlFormatValue = sBetween & _
     sqlFormatValue(pvValue(0)) & sAnd & sqlFormatValue(pvValue(1))
                
                Exit Function
            Else
                Err.Raise 5
            End If
        Case sqlLike
            sqlFormatValue = sLike & sqlFormatValue(pvValue)
            Exit Function
        Case sqlIn
            Err.Raise 5
    End Select
    
    Select Case liType
        Case vbEmpty, vbNull
            Select Case piOperators Mod sqlOR
                Case sqlEqual, sqlNotEqual
                    If piOperators = sqlNotEqual Then sqlFormatValue = sqlFormatValue & sNot
                    sqlFormatValue = sqlFormatValue & " IS NULL"
                Case Else
                    Err.Raise 5
            End Select
        Case vbDate
            If pbSkipValueFormatting Then
                sqlFormatValue = sqlFormatValue & pvValue
            Else
                sqlFormatValue = sqlFormatValue & "#" & pvValue & "#"
            End If
        Case vbString
            If pbSkipValueFormatting Then
                sqlFormatValue = sqlFormatValue & Replace(pvValue, "'", "''")
            Else
                sqlFormatValue = sqlFormatValue & "'" & Replace(Replace(pvValue, "'", "''"), "*", "%") & "'"
            End If
        Case vbObject, vbError, vbDataObject, vbUserDefinedType
            Err.Raise 5
        Case Else
            sqlFormatValue = sqlFormatValue & pvValue
    End Select
End Function

Public Function sqlInOperator(psFieldName As String, pvInObject) _
                 As String
    
    Dim i As Integer
    Dim lvEach As Variant
    Dim loReturn As cStringBuilder
    Set loReturn = New cStringBuilder

    If IsObject(pvInObject) Then
        If Not (TypeOf pvInObject Is Collection Or IsArray(pvInObject)) Then Err.Raise 5
    Else
        If Not IsArray(pvInObject) Then Err.Raise 5
    End If
    
    With loReturn
        .Append psFieldName & sIn & "("
        For Each lvEach In pvInObject
            i = i + 1
            .Append sqlFormatValue(lvEach)
            If i Mod 15 = 0 And i > 0 Then
                .Append ")" & sOr & psFieldName & sIn & "("
            Else
                .Append ", "
            End If
        Next
        If i > 0 Then .Remove .Length - 2, 2
        .Append ")"
    End With
    
    sqlInOperator = loReturn.ToString
End Function

Private Function sqlGetOperator(piOperator As eSQLOperators) _
                 As String
                 
    Select Case piOperator Mod sqlOR
        Case sqlEqual
            sqlGetOperator = sEQL
        Case sqlGreaterThan
            sqlGetOperator = sGT
        Case sqlGreaterThanEqualto
            sqlGetOperator = sGTE
        Case sqlLessThan
            sqlGetOperator = sLT
        Case sqlLessThanEqualTo
            sqlGetOperator = sLTE
        Case sqlNotEqual
            sqlGetOperator = sNotEQL
        Case sqlBetween
            sqlGetOperator = sBetween
    End Select

End Function

Public Function sqlStatementInsert(psTable As String, _
                                psExternalDB As String, _
                          ByVal pCollSubstituteFields As Collection, _
                          ByVal poSELECT As cSELECTStatement) _
                As String
                
    On Error Resume Next
    
    Dim lsField As String
    
    Dim loEach   As cClause
    Dim i        As Long
    Dim liCount  As Long: liCount = poSELECT.SelectClauses.Count
    Dim loReturn As cStringBuilder: Set loReturn = New cStringBuilder
    
    loReturn.Append sInsertInto & psTable & " "
    For Each loEach In poSELECT.SelectClauses
        i = i + 1
        Err.Clear
        With loEach
            If Len(.Text) > 0 Then lsField = .Text Else lsField = .Value
            loReturn.Append pCollSubstituteFields(lsField)
            If Err.Number > 0 Then loReturn.Append lsField
            If i < liCount Then loReturn.Append ", "
        End With
    Next
    If Len(psExternalDB) > 0 Then loReturn.Append sIn & psExternalDB
    loReturn.Append " " & poSELECT.SQLText
    sqlStatementInsert = loReturn.ToString
End Function

Public Function sqlStatementDelete(psTable As String, psWhere As String) As String
    sqlStatementDelete = sDelete & "*" & sFrom & psTable & sWhere & psWhere
End Function

Public Function sqlStatementSelect(psTable As String, _
                                psFrom As String, _
                                psWhere As String) _
                As String
                
    sqlStatementSelect = sSelect & psFrom & sFrom & psTable
    If Len(psWhere) > 0 Then sqlStatementSelect = sqlStatementSelect & sWhere & psWhere
    
End Function

Public Function sqlStatementUpdate(psTable As String, _
                                psSet As String, _
                                psWhere As String) _
                As String
    
    sqlStatementUpdate = sUpdate & psTable & sSet & psSet
    If Len(psWhere) > 0 Then sqlStatementUpdate = sqlStatementUpdate & sWhere & psWhere

End Function

Public Function sqlOrderBy(psField As String, piType As eSQLSortModes) _
                As String
    
    If Len(psField) > 0 Then
        Select Case piType
            Case sqlDescending
                sqlOrderBy = sOrderBy & psField & " DESC"
            Case sqlAscending
                sqlOrderBy = sOrderBy & psField
        End Select
    End If

End Function

Public Function sqlShape(psParent As String, _
                   ByVal pCollChildren As Collection) _
                As String
    
    Dim loEach  As cSELECTStatement
    Dim i       As Long
    Dim liCount As Long: liCount = pCollChildren.Count
    
    Dim loReturn As cStringBuilder: Set loReturn = New cStringBuilder
    loReturn.Append sShape & "{" & psParent & "}" & sAppend

    For Each loEach In pCollChildren
        i = i + 1
        With loEach
            loReturn.Append "({" & .SQLText & "}" & _
                       sRelate & .RelatedOnField & sTo & .RelatedToField & ")" & _
                       sAs & .RelatedAppendFieldName
            If Not i = liCount Then loReturn.Append ", "
        End With
    Next
    sqlShape = loReturn.ToString
End Function

Public Function sqlGroupBy(psGroupBy As String, pcollHaving As Collection) _
                As String
    
    Dim loReturn As cStringBuilder
    
    Dim loEach As cClause
    If pcollHaving.Count > 0 Then
        Set loReturn = New cStringBuilder
        loReturn.Append sGroupBy & psGroupBy & sHaving
        For Each loEach In pcollHaving
            With loEach
                loReturn.Append sqlFieldCompare(.Text, .Value, .Operator, .Parentheses, .Formatted)
            End With
        Next
        sqlGroupBy = loReturn.ToString
        Set loReturn = Nothing
    Else
        sqlGroupBy = sGroupBy & psGroupBy
    End If
End Function



'Non-SQL-------
Public Function BitIsSet(ByVal piVal As Long, ByVal piBit As Long) As Boolean
    BitIsSet = CBool(piVal And piBit)
End Function
Public Sub SetBit(piVal As Long, ByVal piBit As Long, ByVal pbState As Boolean)
    Dim lbVal As Boolean
    lbVal = BitIsSet(piVal, piBit)
    If Not lbVal = pbState Then
        If pbState Then
            piVal = piVal + piBit
        Else
            piVal = piVal - piBit
        End If
    End If
End Sub
'--------------










'Public Sub test()
'    Dim i As Long
'    Dim liTime As Single
'    liTime = Timer
'    Dim loSQL As cSELECTStatement
'    Set loSQL = New cSELECTStatement
'    'Dim loInsert As cINSERTStatement
'    'Set loInsert = New cINSERTStatement
'    'Set loSQL = loInsert.SELECTStatement
'
'    'loInsert.TableName = "NewTable"
'    'loInsert.AddSubstituteField "ASDFField", "Alias5"
'    'loInsert.InsertIntoTable = "InsertIntoTable"
'    With loSQL
'
'        For i = 1 To 5
'            .AddSelectClause "Field" & i, "Alias" & i, i
'        Next
'        .SelectType = sqlDistinctRow
'
'        .AddWhereClause "FieldToMatch1", "MatchValue1", , -1
'        .AddWhereClause "FieldToMatch2", 4567, sqlGreaterThan, 1
'        .AddWhereClause "FieldToMatch3", 9876, sqlGreaterThanEqualto + sqlOR, -1
'        .AddWhereClause "FieldToMatch4", "MatchValue4", sqlLessThan
'        .AddWhereClause "FieldToMatch5", "MatchValue5", sqlLessThanEqualTo
'        .AddWhereClause "FieldToMatch6", "MatchValue6", sqlNotEqual
'        .AddWhereClause "FieldToMatch7", Array(#1/1/1980#, #1/1/1990#, #1/1/2000#), sqlIn
'        .AddWhereClause "FieldToMatch8", Array(#1/1/1980#, #1/1/1990#), sqlBetween
'        .AddWhereClause "FieldToMatch9", "C?D*", sqlLike, 1
'        .TableName = "TableName"
'        .SelectIntoExternalDB = "C:\test.mdb"
'        .SelectIntoTable = "NewTable"
'        .SortColumn = "Field2"
'        .SortMode = sqlDescending
'        .GroupBy = "Field1"
'
'        .AddHavingClause "FieldToMatch1", "MatchValue1", , -1
'        .AddHavingClause "FieldToMatch2", 4567, sqlGreaterThan, 1
'        .AddHavingClause "FieldToMatch3", 9876, sqlGreaterThanEqualto + sqlOR, -1
'        .AddHavingClause "FieldToMatch4", "MatchValue4", sqlLessThan
'        .AddHavingClause "FieldToMatch5", "MatchValue5", sqlLessThanEqualTo
'        .AddHavingClause "FieldToMatch6", "MatchValue6", sqlNotEqual
'        .AddHavingClause "FieldToMatch7", Array(#1/1/1980#, #1/1/1990#, #1/1/2000#), sqlIn
'        .AddHavingClause "FieldToMatch8", Array(#1/1/1980#, #1/1/1990#), sqlBetween
'        .AddHavingClause "FieldToMatch9", "C?D*", sqlLike, 1
'
'        Debug.Print .SQLText
'        'Debug.Print loInsert.SQLText
'    End With
'    Debug.Print CCur(Timer - liTime)
'End Sub
