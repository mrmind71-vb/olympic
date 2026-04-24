VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form SupRpt15 
   Caption         =   " ﬁ«—Ì— «·„Ê—œÌ‰"
   ClientHeight    =   1995
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5235
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
   LinkTopic       =   "Form2"
   RightToLeft     =   -1  'True
   ScaleHeight     =   1995
   ScaleWidth      =   5235
   StartUpPosition =   3  'Windows Default
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   -990
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      RightToLeft     =   -1  'True
      Top             =   810
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.TextBox xCode 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2625
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   225
      Width           =   1290
   End
   Begin VB.TextBox xdate1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2625
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   750
      Visible         =   0   'False
      Width           =   1290
   End
   Begin VB.TextBox xDate2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   150
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   750
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   690
      Left            =   300
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   1200
      Width           =   3690
      Begin VB.CommandButton CmdExit 
         Caption         =   "Œ—ÊÃ"
         Height          =   390
         Left            =   75
         RightToLeft     =   -1  'True
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   165
         Width           =   1065
      End
      Begin VB.CommandButton CmdUndo 
         Caption         =   " —«Ã⁄"
         Height          =   390
         Left            =   1275
         RightToLeft     =   -1  'True
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   165
         Width           =   1065
      End
      Begin VB.CommandButton CmdApply 
         Caption         =   "⁄—÷"
         Height          =   390
         Left            =   2475
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   165
         Width           =   1140
      End
   End
   Begin Crystal.CrystalReport Report1 
      Left            =   4140
      Top             =   1440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTop       =   0
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      BoundReportHeading=   "dddd"
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin VB.Label xCodeName 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   150
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   225
      Width           =   2340
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "„‰  «—ÌŒ :"
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
      Left            =   3990
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   750
      Visible         =   0   'False
      Width           =   765
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "«·Ï :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1425
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   825
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "«·„Ê—œ"
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
      Left            =   4170
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   285
      Width           =   480
   End
End
Attribute VB_Name = "SupRpt15"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SuplierTable As Recordset
Dim TempTable As Recordset
Sub myProc()
ActiveControl.Text = GrdText(Search.Grid1, 0)
Unload Search
End Sub
Private Sub CmdApply_Click()
Dim datatable As Recordset
tempdb.Execute "DELETE * FROM TEMP"
Set TargetTable = tempdb.CreateDynaset("TEMP")

If xCode.Text <> "" Then

    TargetTable.AddNew
    cString = " SELECT sum(FILE7_20.total) AS tinv FROM FILE7_20 where code = " & MyParn(xCode.Text)
    Set datatable = mydb.OpenRecordset(cString, dbOpenDynaset)
    If datatable.RecordCount > 0 Then
        TargetTable.VAL1 = datatable.TINV
    End If
    
    cString = " SELECT sum(FILE6_11.total) AS tinv FROM FILE6_11 where code = " & MyParn(xCode.Text)
    Set datatable = mydb.OpenRecordset(cString, dbOpenDynaset)
    If datatable.RecordCount > 0 Then
        TargetTable.VAL2 = datatable.TINV
    End If
    
    cString = " SELECT sum(FILE8_40.VALUE) AS tinv FROM FILE8_40 where code = " & MyParn(xCode.Text)
    Set datatable = mydb.OpenRecordset(cString, dbOpenDynaset)
    If datatable.RecordCount > 0 Then
        TargetTable.VAL3 = datatable.TINV
    End If
    
    TargetTable.VAL4 = TurnValue(TargetTable.VAL1, Null, 0) - TurnValue(TargetTable.VAL2, Null, 0) - TurnValue(TargetTable.VAL3, Null, 0)
    
    cString = " SELECT Sum(FILE6_20.TOTAL) AS TTOTAL, Sum(FILE6_20.QUANT) AS  TQUANT , Sum(FILE6_20.QUANT * file6_20.cost ) AS  Tcost " & _
            " FROM FILE6_20 LEFT JOIN ITEMCODE ON FILE6_20.ITEM = ITEMCODE.ITEM WHERE ITEMCODE.LastOfCODE =  " & MyParn(xCode.Text)
    Set datatable = mydb.OpenRecordset(cString, dbOpenDynaset)
    
    If datatable.RecordCount > 0 Then
        TargetTable.VAL5 = datatable.Tcost * ((100 - TurnValue(SuplierTable.DISC, Null, 0)) / 100)
        TargetTable.VAL8 = datatable.Tcost * ((100 - TurnValue(SuplierTable.DISC, Null, 0)) / 100)
        TargetTable.VAL7 = datatable.TTOTAL
    End If
    
    cString = " SELECT sum(FILE8_30.VALUE) AS tinv FROM FILE8_30 where code = " & MyParn(xCode.Text)
    Set datatable = mydb.OpenRecordset(cString, dbOpenDynaset)
    If datatable.RecordCount > 0 Then
        TargetTable.VAL9 = datatable.TINV
    End If
    
    cString = " SELECT sum(FILE4_10.F_BAL1 ) AS tinv FROM FILE4_10 where code = " & MyParn(xCode.Text)
    Set datatable = mydb.OpenRecordset(cString, dbOpenDynaset)
    If datatable.RecordCount > 0 Then
        TargetTable.VAL9 = TurnValue(TargetTable.VAL9, Null, 0) + TurnValue(datatable.TINV, Null, 0)
    End If
    
    cString = " SELECT sum(FILE5_21.VALUE) AS tinv FROM FILE5_21 where code = " & MyParn(xCode.Text)
    Set datatable = mydb.OpenRecordset(cString, dbOpenDynaset)
    If datatable.RecordCount > 0 Then
        TargetTable.VAL9 = TurnValue(TargetTable.VAL9, Null, 0) + TurnValue(datatable.TINV, Null, 0)
    End If
    
    TargetTable.VAL10 = TurnValue(TargetTable.VAL8, Null, 0) - TurnValue(TargetTable.VAL9, Null, 0)
    
    
    cStr1 = " SELECT Sum(FILE1_11.OUT * ITEMCODE.cost ) AS VOUT , Sum(FILE1_11.[IN] * ITEMCODE.cost ) AS  VIN  " & _
            " FROM FILE1_11 LEFT JOIN ITEMCODE ON FILE1_11.ITEM = ITEMCODE.ITEM WHERE ITEMCODE.LastOfCODE =  " & MyParn(xCode.Text)
    Set datatable = mydb.OpenRecordset(cStr1, dbOpenDynaset)
    If datatable.RecordCount > 0 Then
      TargetTable.VAL6 = (TurnValue(datatable.VIN, Null, 0) - TurnValue(datatable.VOUT, Null, 0)) * ((100 - TurnValue(SuplierTable.DISC, Null, 0)) / 100)
    End If
    
    TargetTable.STR19 = "„” Õﬁ«  «·„Ê—œ " & xCodeName.Caption
    TargetTable.Update
    myws.BeginTrans
    myws.CommitTrans

    Report1.ReportFileName = PublicPath & "\Reports\R_SUPP.rpt"
    Report1.DataFiles(0) = cPathTemp
    Report1.Action = 1
Else
    SuplierTable.MoveFirst
    Do While Not SuplierTable.EOF
        xCode.Text = SuplierTable.CODE
        xCodeName.Caption = SuplierTable.DESCA
        
        TargetTable.AddNew
        cString = " SELECT sum(FILE7_20.total) AS tinv FROM FILE7_20 where code = " & MyParn(xCode.Text)
        Set datatable = mydb.OpenRecordset(cString, dbOpenDynaset)
        If datatable.RecordCount > 0 Then
            TargetTable.VAL1 = datatable.TINV
        End If
        
        cString = " SELECT sum(FILE6_11.total) AS tinv FROM FILE6_11 where code = " & MyParn(xCode.Text)
        Set datatable = mydb.OpenRecordset(cString, dbOpenDynaset)
        If datatable.RecordCount > 0 Then
            TargetTable.VAL2 = datatable.TINV
        End If
        
        cString = " SELECT sum(FILE8_40.VALUE) AS tinv FROM FILE8_40 where code = " & MyParn(xCode.Text)
        Set datatable = mydb.OpenRecordset(cString, dbOpenDynaset)
        If datatable.RecordCount > 0 Then
            TargetTable.VAL3 = datatable.TINV
        End If
        
        TargetTable.VAL4 = TurnValue(TargetTable.VAL1, Null, 0) - TurnValue(TargetTable.VAL2, Null, 0) - TurnValue(TargetTable.VAL3, Null, 0)
        
        cString = " SELECT Sum(FILE6_20.TOTAL) AS TTOTAL, Sum(FILE6_20.QUANT) AS  TQUANT , Sum(FILE6_20.QUANT * file6_20.cost ) AS  Tcost " & _
                " FROM FILE6_20 LEFT JOIN ITEMCODE ON FILE6_20.ITEM = ITEMCODE.ITEM WHERE ITEMCODE.LastOfCODE =  " & MyParn(xCode.Text)
        Set datatable = mydb.OpenRecordset(cString, dbOpenDynaset)
        
        If datatable.RecordCount > 0 Then
            TargetTable.VAL5 = datatable.Tcost * ((100 - TurnValue(SuplierTable.DISC, Null, 0)) / 100)
            TargetTable.VAL8 = datatable.Tcost * ((100 - TurnValue(SuplierTable.DISC, Null, 0)) / 100)
            TargetTable.VAL7 = datatable.TTOTAL
        End If
        
        cString = " SELECT sum(FILE8_30.VALUE) AS tinv FROM FILE8_30 where code = " & MyParn(xCode.Text)
        Set datatable = mydb.OpenRecordset(cString, dbOpenDynaset)
        If datatable.RecordCount > 0 Then
            TargetTable.VAL9 = datatable.TINV
        End If
        
        cString = " SELECT sum(FILE4_10.F_BAL1 ) AS tinv FROM FILE4_10 where code = " & MyParn(xCode.Text)
        Set datatable = mydb.OpenRecordset(cString, dbOpenDynaset)
        If datatable.RecordCount > 0 Then
            TargetTable.VAL9 = TurnValue(TargetTable.VAL9, Null, 0) + TurnValue(datatable.TINV, Null, 0)
        End If
        
        cString = " SELECT sum(FILE5_21.VALUE) AS tinv FROM FILE5_21 where code = " & MyParn(xCode.Text)
        Set datatable = mydb.OpenRecordset(cString, dbOpenDynaset)
        If datatable.RecordCount > 0 Then
            TargetTable.VAL9 = TurnValue(TargetTable.VAL9, Null, 0) + TurnValue(datatable.TINV, Null, 0)
        End If
        
        TargetTable.VAL10 = TurnValue(TargetTable.VAL8, Null, 0) - TurnValue(TargetTable.VAL9, Null, 0)
        
        
        cStr1 = " SELECT Sum(FILE1_11.OUT * ITEMCODE.cost ) AS VOUT , Sum(FILE1_11.[IN] * ITEMCODE.cost ) AS  VIN  " & _
                " FROM FILE1_11 LEFT JOIN ITEMCODE ON FILE1_11.ITEM = ITEMCODE.ITEM WHERE ITEMCODE.LastOfCODE =  " & MyParn(xCode.Text)
        Set datatable = mydb.OpenRecordset(cStr1, dbOpenDynaset)
        If datatable.RecordCount > 0 Then
          TargetTable.VAL6 = (TurnValue(datatable.VIN, Null, 0) - TurnValue(datatable.VOUT, Null, 0)) * ((100 - TurnValue(SuplierTable.DISC, Null, 0)) / 100)
        End If
        
        TargetTable.STR19 = xCodeName.Caption
        TargetTable.Update
        
        SuplierTable.MoveNext
    Loop
    myws.BeginTrans
    myws.CommitTrans
    Report1.ReportFileName = PublicPath & "\Reports\R_AllSUPP.rpt"
    Report1.DataFiles(0) = cPathTemp
    Report1.Action = 1

End If
End Sub
Private Sub CmdUndo_Click()
xclient.Text = ""
End Sub
Private Sub CmdExit_Click()
Unload Me
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is DBCombo Then SendKeys "{TAB}"
End If
End Sub
Private Sub Form_Load()
Set SuplierTable = mydb.OpenRecordset("File4_10", dbOpenDynaset)
Set TempTable = tempdb.OpenRecordset("TEMP")
Set FlagTable = mydb.OpenRecordset("File1_70")
If publicFlag = 13 Or publicFlag = 14 Then
    Label1.Visible = False
    xCode.Visible = False
    xCodeName.Visible = False
    xdate1.Top = xCode.Top
    xDate2.Top = xCode.Top
    Label3.Top = xCode.Top
    Label4.Top = xCode.Top
End If
End Sub
Function MYVALID() As Boolean
If xCode.Visible Then
    SuplierTable.FindFirst "Code = " & MyParn(xCode.Text)
    If SuplierTable.NoMatch Then Exit Function
End If
If Not (IsDate(xdate1.Text) And IsDate(xDate2.Text)) Then
    MsgBox "«· «—ÌŒ €Ì— ’«·Õ"
    Exit Function
End If
MYVALID = True
End Function

Private Sub xCode_Change()
SuplierTable.FindFirst "Code =" & MyParn(xCode.Text)
xCodeName.Caption = IIf(SuplierTable.NoMatch, "", SuplierTable.DESCA)
End Sub
Private Sub xCode_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
    xCode.Text = ""
    Dim Generalarray(3)
    Dim GrdArray(3)
        
    Set Generalarray(1) = Me
    Generalarray(2) = "Select Code As «·ﬂÊœ,DescA As «·«”„Ê From File4_10"
    Generalarray(3) = "Where DescA Like '*cFilter*'"
        
    GrdArray(1) = 1000
    GrdArray(2) = 2600
    GrdArray(3) = 1500
        
    Lookupdata = Array(Generalarray, GrdArray)
    Load Search
    Search.Caption = "«” ⁄·«„ "
    Search.Show 1
End If
End Sub
