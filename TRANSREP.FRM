VERSION 5.00
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form TransRep 
   Caption         =   " ﬁ—Ì— «· ÕÊÌ·« "
   ClientHeight    =   2685
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5340
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
   ScaleHeight     =   2685
   ScaleWidth      =   5340
   StartUpPosition =   3  'Windows Default
   Begin VB.Data Data2 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   390
      Left            =   -825
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      RightToLeft     =   -1  'True
      Top             =   675
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   390
      Left            =   -960
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      RightToLeft     =   -1  'True
      Top             =   240
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.TextBox xDate2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2850
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   600
      Width           =   1290
   End
   Begin VB.TextBox xdate1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2850
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   150
      Width           =   1290
   End
   Begin Crystal.CrystalReport Report1 
      Left            =   960
      Top             =   225
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
   Begin VB.Frame Frame1 
      Height          =   690
      Left            =   225
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   1875
      Width           =   3765
      Begin VB.CommandButton CmdClear 
         Caption         =   " ÃœÌœ"
         Height          =   390
         Left            =   1260
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   225
         Width           =   1215
      End
      Begin VB.CommandButton CmdExit 
         Caption         =   "Œ—ÊÃ"
         Height          =   390
         Left            =   75
         RightToLeft     =   -1  'True
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   225
         Width           =   1140
      End
      Begin VB.CommandButton CmdApply 
         Caption         =   "⁄—÷"
         Height          =   390
         Left            =   2520
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   225
         Width           =   1140
      End
   End
   Begin MSDBCtls.DBCombo xStore1 
      Bindings        =   "TRANSREP.frx":0000
      Height          =   315
      Left            =   825
      TabIndex        =   2
      Top             =   1050
      Width           =   3315
      _ExtentX        =   5847
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSDBCtls.DBCombo xStore2 
      Bindings        =   "TRANSREP.frx":0014
      Height          =   315
      Left            =   825
      TabIndex        =   3
      Top             =   1500
      Width           =   3315
      _ExtentX        =   5847
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "≈·Ï „Œ“‰ :"
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
      Left            =   4275
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   1575
      Width           =   885
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "„‰ „Œ“‰ :"
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
      Left            =   4275
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   1125
      Width           =   825
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "≈·Ï  «—ÌŒ :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   4275
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   600
      Width           =   765
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
      Height          =   240
      Left            =   4275
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   150
      Width           =   765
   End
End
Attribute VB_Name = "TransRep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TempTable As Recordset
Dim SourceTable As Recordset
Dim nOption As Integer
Function MYVALID()
If Not IsDate(xdate1.Text) Then Exit Function
If Not IsDate(xDate2.Text) Then Exit Function
If xStore1.BoundText = xStore2.BoundText Then
    MsgBox "·« Ì„ﬂ‰ «· ÕÊÌ· „‰ „Œ“‰ «·Ï ‰›” „Œ“‰"
    Exit Function
End If
If xStore1.BoundText = "" And xStore2.BoundText = "" Then Exit Function
MYVALID = True
End Function

Private Sub CmdClear_Click()
xdate1.Text = ""
xDate2.Text = ""
xStore1.BoundText = ""
xStore2.BoundText = ""
End Sub

Private Sub CmdExit_Click()
Unload Me
End Sub
Private Sub CmdUndo_Click()
xStore1.BoundText = ""
xdate1.Text = ""
xDate2.Text = ""
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is DBCombo Then SendKeys "{TAB}"
End If
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 46 Then
    If TypeOf ActiveControl Is DBCombo Then ActiveControl.BoundText = ""
End If
End Sub

Private Sub Form_Load()
Set TempTable = tempdb.OpenRecordset("Temp")
Data1.DatabaseName = MdbPath
Data1.RecordSource = "SELECT CODE , DESCA FROM FILE1_70 WHERE FLAG = 1 "
Data2.DatabaseName = MdbPath
Data2.RecordSource = "SELECT CODE , DESCA FROM FILE1_70 WHERE FLAG = 1 "
xStore1.ListField = "Desca"
xStore1.BoundColumn = "code"
xStore2.ListField = "Desca"
xStore2.BoundColumn = "code"
End Sub
Private Sub CmdApply_Click()
If Not MYVALID Then Exit Sub
Select Case publicFlag
Case 10
    RepItem1
End Select
End Sub
Sub RepItem1()
tempdb.Execute "DELETE * FROM TEMP"
cString = "Select File1_60.Item as MidOfItem," & _
          "Sum(File1_60.Quant) As SumOfAmount," & _
          "First(File1_10.DescA) as FirstOfDescA" & _
          " From File1_60 inner join File1_10 on File1_60.item = file1_10.item " & _
          " where Date Between DateValue(" & MyParn(xdate1.Text) & ")" & _
          " and DateValue(" & MyParn(xDate2.Text) & ")" & _
          IIf(xStore1.Text <> "", " and Store1 = " & MyParn(xStore1.BoundText), "") & _
          IIf(xStore2.Text <> "", " and Store2 = " & MyParn(xStore2.BoundText), "") & _
          " Group By " & _
          "File1_60.item," & _
          "File1_10.DescA"
          
Set SourceTable = mydb.OpenRecordset(cString, dbOpenSnapshot)
With SourceTable
If .RecordCount > 0 Then
    Do While Not .EOF
            TempTable.AddNew
            TempTable.str1 = .MidofItem
            TempTable.str2 = .FirstofDescA
            TempTable.str3 = .SumOfAmount
            TempTable.str4 = IIf(xStore1.BoundText <> "", " „‰ " & xStore1.Text, "") & _
                             IIf(xStore2.BoundText <> "", " ≈·Ï " & xStore2.Text, "")
            
            TempTable.Date1 = xdate1.Text
            TempTable.date2 = xDate2.Text
            TempTable.str7 = " ≈Ã„«·Ï  ÕÊÌ·«  „‰ „Œ“‰ " & xStore1.Text & " ≈·Ï „Œ“‰ " & xStore2.Text
            TempTable.str8 = " „‰ " & xdate1.Text & " ≈·Ï " & xDate2.Text
            TempTable.str9 = firsttitle
            TempTable.str10 = Secondtitle
            TempTable.Update
            .MoveNext
    Loop
End If
End With
myws.BeginTrans
myws.CommitTrans
Report1.ReportFileName = PublicPath & "\Reports\TransRep.rpt"
Report1.DataFiles(0) = cPathTemp
Report1.Action = 1
End Sub

Sub temp()
If publicFlag = 1 Then
    RepItem1
Else
    cString = "SELECT MID(FILE1_60.ITEM,1,6) AS UItem ,FIRST(FILE1_10.DESCA) AS DescItem , FIRST(FILE1_10.PACK) AS PackItem , " & _
              " SUM(amount1) AS T_Quant1 , " & _
              " SUM(amount2) AS T_Quant2 , " & _
              " SUM(FILE1_10.COST*(FILE1_60.AMOUNT1+(FILE1_60.AMOUNT2/FILE1_10.PACK))) AS T_Value " & _
              " From File1_60 Inner Join file1_10 on File1_60.Item = File1_10.Item " & _
              " Where vDate >= DateValue(" & MyParn(xdate1.Text) & ")" & _
              " and vDate <= DateValue(" & MyParn(xDate2.Text) & ")" & _
              " And File1_60.NO2 = " & MyParn(xStore1.BoundText) & _
              " Group by Mid(File1_60.Item,1,6) "
End If
Set SourceTable = mydb.CreateSnapshot(cString)
tempdb.Execute "DELETE * FROM TEMP"
cHead = xStore1.Text
With SourceTable
If .RecordCount > 0 Then
.MoveFirst
nTotal = 0
Do While Not .EOF
    N1 = QuantToUnits(.T_Quant1, .T_Quant2, .PACKITEM)
    TempTable.AddNew
    TempTable.str1 = .uitem
    TempTable.str2 = .DESCITEM
    TempTable.str7 = cHead
    TempTable.str8 = " „‰  «—ÌŒ: " & xdate1.Text & " ≈·Ï  «—ÌŒ " & xDate2.Text
    TempTable.str3 = N1
    TempTable.VAL1 = .T_Value
    TempTable.STR19 = firsttitle
    TempTable.str20 = Secondtitle

    TempTable.Update
    .MoveNext
Loop
End If
If TempTable.RecordCount = 0 Then
    MsgBox "·«  ÊÃœ »Ì«‰«  ›Ï «· ﬁ—Ì— ·⁄—÷Â«"
    Exit Sub
End If

End With
myws.BeginTrans
myws.CommitTrans
Report1.ReportFileName = PublicPath & "\Reports\TRANSREP.rpt"
Report1.DataFiles(0) = cPathTemp
Report1.Action = 1
End Sub
