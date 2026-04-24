VERSION 5.00
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form R_Item17 
   Caption         =   " ﬁ«—Ì— "
   ClientHeight    =   2580
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6000
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   RightToLeft     =   -1  'True
   ScaleHeight     =   2580
   ScaleWidth      =   6000
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox xDate 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3075
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   1350
      Width           =   1515
   End
   Begin VB.CommandButton CmdApply2 
      Caption         =   "⁄—÷ ﬁÌ„… «·√—’œ…"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   2100
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   1950
      Width           =   1665
   End
   Begin VB.Data Data3 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   600
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      RightToLeft     =   -1  'True
      Top             =   0
      Visible         =   0   'False
      Width           =   1440
   End
   Begin MSDBCtls.DBCombo xGroup1 
      Bindings        =   "R_Item17.frx":0000
      DataSource      =   "Data1"
      Height          =   315
      Left            =   2550
      TabIndex        =   7
      Top             =   525
      Width           =   2040
      _ExtentX        =   3598
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin VB.CommandButton CmdApply 
      Caption         =   "⁄—÷ √—’œ…"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   4125
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   1950
      Width           =   1665
   End
   Begin VB.CommandButton CmdExit 
      Caption         =   "Œ—ÊÃ"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   75
      RightToLeft     =   -1  'True
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1950
      Width           =   1665
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   450
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      RightToLeft     =   -1  'True
      Top             =   225
      Visible         =   0   'False
      Width           =   1440
   End
   Begin MSDBCtls.DBCombo xGroup 
      Bindings        =   "R_Item17.frx":0014
      DataSource      =   "Data3"
      Height          =   315
      Left            =   2520
      TabIndex        =   0
      Top             =   150
      Width           =   2040
      _ExtentX        =   3598
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin Crystal.CrystalReport REPORT1 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      WindowTop       =   0
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      BoundReportHeading=   "dddd"
      WindowState     =   2
   End
   Begin MSDBCtls.DBCombo xGroup2 
      Bindings        =   "R_Item17.frx":0028
      DataSource      =   "Data1"
      Height          =   315
      Left            =   2550
      TabIndex        =   1
      Top             =   900
      Width           =   2040
      _ExtentX        =   3598
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   " «—ÌŒ"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   2
      Left            =   4725
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   1350
      Width           =   390
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000B&
      Caption         =   "„‰ „Ã„Ê⁄… "
      Height          =   315
      Index           =   0
      Left            =   4725
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   525
      Width           =   1005
   End
   Begin VB.Label Label3 
      BackColor       =   &H8000000B&
      Caption         =   "≈·Ï „Ã„Ê⁄…"
      Height          =   255
      Left            =   4725
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   900
      Width           =   1005
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000B&
      Caption         =   "„Ã„Ê⁄… —∆Ì”Ì…"
      Height          =   195
      Left            =   4725
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   225
      Width           =   1110
   End
End
Attribute VB_Name = "R_Item17"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim BalTable As Recordset
Dim itemTable As Recordset
Dim storeTable As Recordset
Dim CodeTable As Recordset
Private Sub CmdApply2_Click()
tempdb.Execute "DELETE * FROM TEMP"
Dim TempTable As Recordset
Dim nCost As Double
Set TempTable = tempdb.OpenRecordset("Temp")
Set CodeTable = mydb.OpenRecordset("SELECT * FROM FILE1_70 ")
Dim LastCost As Recordset

If IsDate(xDate.Text) Then
    cStr1 = " SELECT FILE7_20.ITEM, FILE7_20.Price, FILE7_20.DATE FROM FILE7_20  " & _
            " WHERE Date <= DateValue(" & MyParn(xDate.Text) & ")" & _
            " ORDER BY DATE "
    Set LastCost = mydb.OpenRecordset(cStr1)
Else
    cStr1 = " SELECT FILE7_20.ITEM, FILE7_20.Price, FILE7_20.DATE FROM FILE7_20  " & _
            " ORDER BY DATE "
    Set LastCost = mydb.OpenRecordset(cStr1)
End If


cString = "SELECT FILE1_11.ITEM, FILE1_11.STORE, Sum(FILE1_11.[IN]) AS SumIN, Sum(FILE1_11.OUT) AS SumOUT  " & _
          " FROM FILE1_11 GROUP BY FILE1_11.ITEM, FILE1_11.STORE "

Set BalTable = mydb.OpenRecordset(cString, dbOpenSnapshot)

cString = "SELECT Sum(FILE1_11.[IN]) AS SumIN, Sum(FILE1_11.OUT) AS SumOUT, FILE1_11.ITEM,  " & _
          " FILE1_10.DESCA , FILE1_10.COST , " & _
          " First(FILE1_50.M_GROUP) AS MGrCode, First(FILE1_50.DESCA) AS GrDESC,  " & _
          " First(FILE1_10.GROUP) AS GrCode, First(FILE1_70.DESCA) AS MGrDesc" & _
          " FROM ((FILE1_10 RIGHT JOIN FILE1_11 ON FILE1_10.ITEM = FILE1_11.ITEM) LEFT " & _
          " JOIN FILE1_50 ON FILE1_10.GROUP = FILE1_50.CODE) LEFT JOIN FILE1_70 ON " & _
          " FILE1_50.M_GROUP = FILE1_70.CODE" & _
          " where file1_70.FLAG = 2 "

If xGroup.BoundText <> "" Then
    cString = cString & " AND file1_50.M_GROUP = " & MyParn(xGroup.BoundText)
End If
    If xGroup1.Text <> "" Then cString = cString & " AND file1_10.GROUP >= " & MyParn(xGroup1.BoundText)
    If xGroup2.Text <> "" Then cString = cString & " AND file1_10.GROUP <= " & MyParn(xGroup2.BoundText)


If IsDate(xDate.Text) Then cString = cString & " and Date <= DateValue(" & MyParn(xDate.Text) & ")"

cString = cString & " GROUP BY FILE1_11.ITEM,  FILE1_10.DESCA ,FILE1_10.COST  "
Set itemTable = mydb.OpenRecordset(cString, dbOpenSnapshot)

cString = "SELECT FILE1_11.STORE  " & _
          " FROM ((FILE1_10 RIGHT JOIN FILE1_11 ON FILE1_10.ITEM = FILE1_11.ITEM) LEFT JOIN FILE1_50 ON  " & _
          " FILE1_10.GROUP = FILE1_50.CODE) LEFT JOIN FILE1_70 ON FILE1_50.M_GROUP = FILE1_70.CODE " & _
          " where file1_70.FLAG = 2 "

If xGroup1.BoundText <> "" Then
    cString = cString & " AND file1_50.M_GROUP = " & MyParn(xGroup.BoundText)
    If xGroup1.Text <> "" Then cString = cString & " AND file1_10.GROUP >= " & MyParn(xGroup1.BoundText)
    If xGroup2.Text <> "" Then cString = cString & " AND file1_10.GROUP <= " & MyParn(xGroup2.BoundText)
End If
If IsDate(xDate.Text) Then cString = cString & " and Date <= DateValue(" & MyParn(xDate.Text) & ")"

cString = cString & " GROUP BY FILE1_11.STORE "
Set storeTable = mydb.OpenRecordset(cString, dbOpenSnapshot)
 
I = 1
With itemTable
If .RecordCount > 0 Then
    Do While Not .EOF
        nBal = TurnValue(.SumIn, Null, 0) - TurnValue(.SUMout, Null, 0)
        If nBal <> 0 Then
            nCost = 0
            If IsDate(xDate.Text) Then
                LastCost.FindLast " ITEM = " & MyParn(.Item)
                If LastCost.NoMatch Then
                    nCost = .cost
                Else
                    nCost = LastCost.price
                End If
            Else
                nCost = .cost
            End If
            
            TempTable.AddNew
            TempTable.str6 = .MGrdesc
            TempTable.STR5 = .MGrCode
            TempTable.str1 = .Item
            TempTable.str2 = .DESCA
            TempTable.str3 = .GrCode
            TempTable.str4 = .GRDESC
            TempTable.val2 = nBal * nCost
            storeTable.MoveFirst
            
            I = 0
            Do While True
                I = I + 1
                Select Case I
                    Case 1
                        TempTable.Str11 = Say2Code(CodeTable, 1, storeTable.Store)
                    Case 2
                        TempTable.Str12 = Say2Code(CodeTable, 1, storeTable.Store)
                    Case 3
                        TempTable.Str13 = Say2Code(CodeTable, 1, storeTable.Store)
                    Case 4
                        TempTable.Str14 = Say2Code(CodeTable, 1, storeTable.Store)
                    Case 5
                        TempTable.str15 = Say2Code(CodeTable, 1, storeTable.Store)
                    Case 6
                        TempTable.str16 = Say2Code(CodeTable, 1, storeTable.Store)
                    Case 7
                        TempTable.Str17 = Say2Code(CodeTable, 1, storeTable.Store)
                    Case 8
                        TempTable.str18 = Say2Code(CodeTable, 1, storeTable.Store)
                End Select
                
                BalTable.FindFirst " ITEM = " & MyParn(.Item) & " AND STORE = " & MyParn(storeTable.Store)
                If Not BalTable.NoMatch Then
                    Select Case I
                        Case 1
                            TempTable.VAL11 = (TurnValue(BalTable.SumIn, Null, 0) - TurnValue(BalTable.SUMout, Null, 0)) * nCost
                        Case 2
                            TempTable.VAL12 = (TurnValue(BalTable.SumIn, Null, 0) - TurnValue(BalTable.SUMout, Null, 0)) * nCost
                        Case 3
                            TempTable.VAL13 = (TurnValue(BalTable.SumIn, Null, 0) - TurnValue(BalTable.SUMout, Null, 0)) * nCost
                        Case 4
                            TempTable.VAL14 = (TurnValue(BalTable.SumIn, Null, 0) - TurnValue(BalTable.SUMout, Null, 0)) * nCost
                        Case 5
                            TempTable.VAL15 = (TurnValue(BalTable.SumIn, Null, 0) - TurnValue(BalTable.SUMout, Null, 0)) * nCost
                        Case 6
                            TempTable.VAL16 = (TurnValue(BalTable.SumIn, Null, 0) - TurnValue(BalTable.SUMout, Null, 0)) * nCost
                        Case 7
                            TempTable.VAL17 = (TurnValue(BalTable.SumIn, Null, 0) - TurnValue(BalTable.SUMout, Null, 0)) * nCost
                        Case 8
                            TempTable.VAL18 = (TurnValue(BalTable.SumIn, Null, 0) - TurnValue(BalTable.SUMout, Null, 0)) * nCost
                    End Select
                End If
                storeTable.MoveNext
                If storeTable.EOF Then Exit Do
            Loop
            TempTable.str7 = "  Ê“Ì⁄ ﬁÌ„… «·√—’œ… ⁄·Ï «·„Œ«“‰ "
            TempTable.str8 = firsttitle
            TempTable.str9 = Secondtitle
            TempTable.Update
        End If
       .MoveNext
    Loop
End If
End With
Report1.ReportFileName = PublicPath & "\Reports\R_item_17.RPt"
Report1.DataFiles(0) = App.Path & "\Temp.mdb"
Report1.Action = 1
End Sub
Private Sub CmdExit_Click()
Unload Me
End Sub
Private Sub Form_Load()
Data1.DatabaseName = MdbPath
Data1.RecordSource = "File1_50 "
xGroup1.BoundColumn = "Code"
xGroup1.ListField = "Desca"

xGroup2.BoundColumn = "Code"
xGroup2.ListField = "Desca"

Data3.DatabaseName = MdbPath
Data3.RecordSource = "Select * From File1_70 Where flag = 2"
xGroup.BoundColumn = "Code"
xGroup.ListField = "Desca"

End Sub
Private Function MYVALID()
MYVALID = True
End Function
Private Sub xGroup_Change()
Data1.RecordSource = "Select * From File1_50 where M_GROUP = " & MyParn(xGroup.BoundText)
Data1.Refresh
End Sub
Private Sub CmdApply_Click()
tempdb.Execute "DELETE * FROM TEMP"
Dim TempTable As Recordset

Set TempTable = tempdb.OpenRecordset("Temp")
Set CodeTable = mydb.OpenRecordset("SELECT * FROM FILE1_70 ")

cString = "SELECT FILE1_11.ITEM, FILE1_11.STORE, Sum(FILE1_11.[IN]) AS SumIN, Sum(FILE1_11.OUT) AS SumOUT  " & _
          " FROM FILE1_11 GROUP BY FILE1_11.ITEM, FILE1_11.STORE "

Set BalTable = mydb.OpenRecordset(cString, dbOpenSnapshot)

cString = "SELECT Sum(FILE1_11.[IN]) AS SumIN, Sum(FILE1_11.OUT) AS SumOUT, FILE1_11.ITEM,  " & _
          " FILE1_10.DESCA , " & _
          " First(FILE1_50.M_GROUP) AS MGrCode, First(FILE1_50.DESCA) AS GrDESC,  " & _
          " First(FILE1_10.GROUP) AS GrCode, First(FILE1_70.DESCA) AS MGrDesc" & _
          " FROM ((FILE1_10 RIGHT JOIN FILE1_11 ON FILE1_10.ITEM = FILE1_11.ITEM) LEFT " & _
          " JOIN FILE1_50 ON FILE1_10.GROUP = FILE1_50.CODE) LEFT JOIN FILE1_70 ON " & _
          " FILE1_50.M_GROUP = FILE1_70.CODE" & _
          " where file1_70.FLAG = 2 "

If xGroup.BoundText <> "" Then
    cString = cString & " AND file1_50.M_GROUP = " & MyParn(xGroup.BoundText)
End If
If xGroup1.Text <> "" Then cString = cString & " AND file1_10.GROUP >= " & MyParn(xGroup1.BoundText)
If xGroup2.Text <> "" Then cString = cString & " AND file1_10.GROUP <= " & MyParn(xGroup2.BoundText)


If IsDate(xDate.Text) Then cString = cString & " and Date <= DateValue(" & MyParn(xDate.Text) & ")"

cString = cString & " GROUP BY FILE1_11.ITEM,  FILE1_10.DESCA  "
Set itemTable = mydb.OpenRecordset(cString, dbOpenSnapshot)

cString = "SELECT FILE1_11.STORE  " & _
          " FROM ((FILE1_10 RIGHT JOIN FILE1_11 ON FILE1_10.ITEM = FILE1_11.ITEM) LEFT JOIN FILE1_50 ON  " & _
          " FILE1_10.GROUP = FILE1_50.CODE) LEFT JOIN FILE1_70 ON FILE1_50.M_GROUP = FILE1_70.CODE " & _
          " where file1_70.FLAG = 2 "

If xGroup1.BoundText <> "" Then
    cString = cString & " AND file1_50.M_GROUP = " & MyParn(xGroup.BoundText)
    If xGroup1.Text <> "" Then cString = cString & " AND file1_10.GROUP >= " & MyParn(xGroup1.BoundText)
    If xGroup2.Text <> "" Then cString = cString & " AND file1_10.GROUP <= " & MyParn(xGroup2.BoundText)
End If
If IsDate(xDate.Text) Then cString = cString & " and Date <= DateValue(" & MyParn(xDate.Text) & ")"

cString = cString & " GROUP BY FILE1_11.STORE "
Set storeTable = mydb.OpenRecordset(cString, dbOpenSnapshot)
 
I = 1
With itemTable
If .RecordCount > 0 Then
    Do While Not .EOF
        nBal = TurnValue(.SumIn, Null, 0) - TurnValue(.SUMout, Null, 0)
        If nBal <> 0 Then
            TempTable.AddNew
            TempTable.str6 = .MGrdesc
            TempTable.STR5 = .MGrCode
            TempTable.str1 = .Item
            TempTable.str2 = .DESCA
            TempTable.str3 = .GrCode
            TempTable.str4 = .GRDESC
            TempTable.val2 = nBal
            storeTable.MoveFirst
            
            I = 0
            Do While True
                I = I + 1
                Select Case I
                    Case 1
                        TempTable.Str11 = Say2Code(CodeTable, 1, storeTable.Store)
                    Case 2
                        TempTable.Str12 = Say2Code(CodeTable, 1, storeTable.Store)
                    Case 3
                        TempTable.Str13 = Say2Code(CodeTable, 1, storeTable.Store)
                    Case 4
                        TempTable.Str14 = Say2Code(CodeTable, 1, storeTable.Store)
                    Case 5
                        TempTable.str15 = Say2Code(CodeTable, 1, storeTable.Store)
                    Case 6
                        TempTable.str16 = Say2Code(CodeTable, 1, storeTable.Store)
                    Case 7
                        TempTable.Str17 = Say2Code(CodeTable, 1, storeTable.Store)
                    Case 8
                        TempTable.str18 = Say2Code(CodeTable, 1, storeTable.Store)
                End Select
                
                BalTable.FindFirst " ITEM = " & MyParn(.Item) & " AND STORE = " & MyParn(storeTable.Store)
                If Not BalTable.NoMatch Then
                    Select Case I
                        Case 1
                            TempTable.VAL11 = TurnValue(BalTable.SumIn, Null, 0) - TurnValue(BalTable.SUMout, Null, 0)
                        Case 2
                            TempTable.VAL12 = TurnValue(BalTable.SumIn, Null, 0) - TurnValue(BalTable.SUMout, Null, 0)
                        Case 3
                            TempTable.VAL13 = TurnValue(BalTable.SumIn, Null, 0) - TurnValue(BalTable.SUMout, Null, 0)
                        Case 4
                            TempTable.VAL14 = TurnValue(BalTable.SumIn, Null, 0) - TurnValue(BalTable.SUMout, Null, 0)
                        Case 5
                            TempTable.VAL15 = TurnValue(BalTable.SumIn, Null, 0) - TurnValue(BalTable.SUMout, Null, 0)
                        Case 6
                            TempTable.VAL16 = TurnValue(BalTable.SumIn, Null, 0) - TurnValue(BalTable.SUMout, Null, 0)
                        Case 7
                            TempTable.VAL17 = TurnValue(BalTable.SumIn, Null, 0) - TurnValue(BalTable.SUMout, Null, 0)
                        Case 8
                            TempTable.VAL18 = TurnValue(BalTable.SumIn, Null, 0) - TurnValue(BalTable.SUMout, Null, 0)
                    End Select
                End If
                storeTable.MoveNext
                If storeTable.EOF Then Exit Do
            Loop
            TempTable.str7 = "  Ê“Ì⁄ «·√—’œ… ⁄·Ï «·„Œ«“‰ "
            TempTable.str8 = firsttitle
            TempTable.str9 = Secondtitle
            TempTable.Update
        End If
       .MoveNext
    Loop
End If
End With
Report1.ReportFileName = PublicPath & "\Reports\R_item_17.RPt"
Report1.DataFiles(0) = App.Path & "\Temp.mdb"
Report1.Action = 1
End Sub


