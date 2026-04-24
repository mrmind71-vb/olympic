VERSION 5.00
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form Chart21 
   BackColor       =   &H00FFFFFF&
   Caption         =   "»Ì«‰Ï"
   ClientHeight    =   5580
   ClientLeft      =   165
   ClientTop       =   360
   ClientWidth     =   10245
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
   ScaleHeight     =   5580
   ScaleWidth      =   10245
   Begin VB.TextBox xDate2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   5925
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   75
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox xDate1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   8475
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   75
      Visible         =   0   'False
      Width           =   1455
   End
   Begin MSChart20Lib.MSChart Chart1 
      Height          =   4365
      Left            =   75
      OleObjectBlob   =   "Chart21.frx":0000
      TabIndex        =   1
      Top             =   450
      Width           =   9690
   End
   Begin VB.CommandButton CmdExit 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   150
      MaskColor       =   &H00FFFFFF&
      Picture         =   "Chart21.frx":1A85
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   0
      TabStop         =   0   'False
      ToolTipText     =   "Œ—ÊÃ"
      Top             =   4875
      UseMaskColor    =   -1  'True
      Width           =   990
   End
End
Attribute VB_Name = "Chart21"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim nOption As Byte, cardtable As Recordset, cField As String
Dim cString1 As String
Dim cString2 As String
Dim nRecordCount As Double
Dim nRec As Double
Private Sub CmdApply_Click()
If cardtable.RecordCount = 0 Then
    MsgBox "·«  ÊÃœ »Ì«‰«  ·⁄—÷Â«"
    EmptyProc
Else
'  ApplyProc
End If
End Sub
Private Sub CmdExit_Click()
Unload Me
End Sub
Private Sub FillChart()
Dim arrValues()
Dim I As Integer
nRec = cardtable.RecordCount
ReDim arrValues(1 To nRec, 1 To 3)
cardtable.MoveFirst
For I = 1 To nRec
    arrValues(I, 1) = cardtable.DESCA
    arrValues(I, 2) = cardtable.SUMTOTAL
    arrValues(I, 3) = cardtable.SUMTOTAL - cardtable.SUMCOST
    cardtable.MoveNext
Next I
Chart1.ChartData = arrValues
Chart1.Visible = True
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is DBCombo Then SendKeys "{tAB}"
End If
End Sub
Private Sub Form_Load()
Me.Picture = LoadPicture(Photopath)
xDate1.Text = R_Item21.Date1.Text
xDate2.Text = R_Item21.date2.Text
cString = "SELECT FILE1_70.CODE, FILE1_70.DESCA, Sum(FILE6_20.TOTAL) AS SumTOTAL, Sum([QUANT]*[COST]) AS SumCost " & _
          "FROM FILE1_70 LEFT JOIN FILE6_20 ON FILE1_70.CODE = FILE6_20.STORE " & _
          "WHERE FILE1_70.FLAG = 1 " & _
          "AND Date >= DateValue(" & MyParn(xDate1.Text) & ")" & _
          "and Date <= DateValue(" & MyParn(xDate2.Text) & ")" & _
          "GROUP BY FILE1_70.CODE, FILE1_70.DESCA "
Set cardtable = mydb.OpenRecordset(cString, dbOpenSnapshot)
FillChart
End Sub
Private Sub Option_Click(Index As Integer, Value As Integer)
nOption = Index
End Sub
Private Function MYVALID()
If Not IsDate(xDate1.Text) Or Not IsDate(xDate2.Text) Then
    MsgBox "Â‰«ﬂ Œÿ√ ›Ï ≈œŒ«· «· «—ÌŒ "
    xDate1.SetFocus
    Exit Function
End If
If DateValue(xDate1.Text) > DateValue(xDate2.Text) Then
    MsgBox "«· «—ÌŒ «·√Ê· √ﬂ»— „‰ «· «—ÌŒ «·À«‰Ï"
    xDate1.SetFocus
    Exit Function
End If

MYVALID = True
End Function
Sub EmptyProc()
Me.Height = 3500
Chart1.Visible = False
CmdExit.Top = 1680
End Sub
Sub ApplyProc()
Me.Height = 7800
Chart1.Visible = True
CmdExit.Top = 6500
FillChart
End Sub
Private Sub xPrice_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tAB}"
End Sub
