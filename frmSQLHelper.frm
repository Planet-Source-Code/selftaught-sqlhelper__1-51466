VERSION 5.00
Begin VB.Form frmSQLHelper 
   Caption         =   "SQLHelper Demo"
   ClientHeight    =   5760
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   7815
   LinkTopic       =   "Form1"
   ScaleHeight     =   5760
   ScaleWidth      =   7815
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txt 
      Height          =   795
      Index           =   1
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   8
      Top             =   1539
      Width           =   7515
   End
   Begin VB.TextBox txt 
      Height          =   795
      Index           =   4
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      Top             =   4860
      Width           =   7515
   End
   Begin VB.TextBox txt 
      Height          =   795
      Index           =   3
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Top             =   3751
      Width           =   7515
   End
   Begin VB.TextBox txt 
      Height          =   795
      Index           =   2
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   2645
      Width           =   7515
   End
   Begin VB.TextBox txt 
      Height          =   795
      Index           =   0
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   433
      Width           =   7515
   End
   Begin VB.Label lbl 
      Caption         =   "Select Statement (Shaped):"
      Height          =   195
      Index           =   4
      Left            =   120
      TabIndex        =   9
      Top             =   1290
      Width           =   1995
   End
   Begin VB.Label lbl 
      Caption         =   "Insert Statement:"
      Height          =   195
      Index           =   3
      Left            =   120
      TabIndex        =   6
      Top             =   4604
      Width           =   1455
   End
   Begin VB.Label lbl 
      Caption         =   "Delete Statement:"
      Height          =   195
      Index           =   2
      Left            =   120
      TabIndex        =   4
      Top             =   3498
      Width           =   1455
   End
   Begin VB.Label lbl 
      Caption         =   "Update Statement:"
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   2392
      Width           =   1455
   End
   Begin VB.Label lbl 
      Caption         =   "Select Statement:"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   180
      Width           =   1455
   End
End
Attribute VB_Name = "frmSQLHelper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Dim loSelect As cSELECTStatement: Set loSelect = New cSELECTStatement
    Dim loUpdate As cUPDATEStatement: Set loUpdate = New cUPDATEStatement
    Dim loDelete As cDELETEStatement: Set loDelete = New cDELETEStatement
    Dim loInsert As cINSERTStatement: Set loInsert = New cINSERTStatement
    
    Dim liTimer As Single
    
    liTimer = Timer
    
    With loSelect
        .AddSelectClause "FieldName1", "Alias1"
        .AddSelectClause "FieldName2", "Alias2"
        .AddSelectClause "AveragedColumn", "ColAverage", sqlAvg
        .AddSelectClause "LastName & ', ' & FirstName", "FullName"
        
        .AddWhereClause "FieldName1", #12/1/2002#, sqlEqual, -1
        .AddWhereClause "FieldName2", "ThisString", sqlNotEqual + sqlAND, 1
        .AddWhereClause "FieldName3", Array(100, 500), sqlBetween + sqlOR, -1
        .AddWhereClause "FieldName4", Array("Val1", "Val2", "Val3", "Val4", "Val5", "Val6", "Val7"), sqlIn + sqlOR, 1
        .AddWhereClause "FieldName5", "FieldName6 / 2", sqlGreaterThan + sqlOR, , True
        
        .TableName = "TableName"
        .GroupBy = "Field1"
        .AddHavingClause "FieldName1", #12/1/2002#, sqlEqual, -1
        .AddHavingClause "FieldName2", "ThisString", sqlNotEqual + sqlAND, 1
        .AddHavingClause "FieldName3", Array(100, 500), sqlBetween + sqlOR, -1
        .AddHavingClause "FieldName4", Array("Val1", "Val2", "Val3", "Val4", "Val5", "Val6", "Val7"), sqlIn + sqlOR, 1
        
        txt(0).Text = .SQLText
        
        .AddChild "ChildTable", "RelatedOnField", "RelatedToField", "RSChild"
        .AddChild "ChildTable2", "RelatedOnField", "RelatedToField", "RSChild2"
        txt(1).Text = .SQLText
    End With
    
    With loUpdate
        .AddSetClause "FieldName1", "Value1"
        .AddSetClause "FieldName2", "FieldName2 * 3", True
        .AddSetClause "FieldName3", 123
        .AddSetClause "FieldName4", #12/1/2002#
        
        .AddWhereClause "FieldName1", #12/1/2002#, sqlEqual, -1
        .AddWhereClause "FieldName2", "ThisString", sqlNotEqual + sqlAND, 1
        .AddWhereClause "FieldName3", Array(100, 500), sqlBetween + sqlOR, -1
        .AddWhereClause "FieldName4", Array("Val1", "Val2", "Val3", "Val4", "Val5", "Val6", "Val7"), sqlIn + sqlOR, 1
        .AddWhereClause "FieldName5", "FieldName6 / 2", sqlGreaterThan + sqlOR, , True
        
        .TableName = "TableName"
        txt(2).Text = .SQLText
    End With
    
    With loDelete
        .TableName = "TableName"
        .AddWhereClause "FieldName1", #12/1/2002#, sqlEqual, -1
        .AddWhereClause "FieldName2", "ThisString", sqlNotEqual + sqlAND, 1
        .AddWhereClause "FieldName3", Array(100, 500), sqlBetween + sqlOR, -1
        .AddWhereClause "FieldName4", Array("Val1", "Val2", "Val3", "Val4", "Val5", "Val6", "Val7"), sqlIn + sqlOR, 1
        .AddWhereClause "FieldName5", "FieldName6 / 2", sqlGreaterThan + sqlOR, , True
        txt(3).Text = .SQLText
    End With
    
    With loInsert
        .AddSubstituteField "NewFieldDifferentName", "FromField"
        .InsertIntoTable = "C:\Temp\ExternalDB.mdb"
        .TableName = "TableName"
        
        With .SELECTStatement
            .AddSelectClause "FromField"
            .AddSelectClause "SecondFromField"
            .AddWhereClause "ReadyForArchive", True
        End With
        txt(4).Text = .SQLText
    
    End With
    
    liTimer = Timer - liTimer
    
    Dim liLen As Long
    
    With txt
        liLen = Len(.Item(0).Text) + Len(.Item(1).Text) + Len(.Item(2).Text) + Len(.Item(3).Text) + Len(.Item(4).Text)
    End With
    Show
    Dim lsglTime As Single
    lsglTime = liTimer / 1000
    If lsglTime <= 0 Then lsglTime = 1.401298E-45 'Must avoid division by zero (yes, it happened!)
    MsgBox "Created 5 statements totaling " & liLen & " characters in " & lsglTime & " seconds." & vbCrLf & vbCrLf & "That means that this could be done " & CDbl(1 / lsglTime) & " times per second."
    
End Sub
