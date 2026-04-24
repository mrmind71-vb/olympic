VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form TSalItem2 
   BackColor       =   &H00FFFFFF&
   Caption         =   "√” «– «·√’‰«› - ≈Ã„«·Ï „»Ì⁄«  «·„Ã„Ê⁄«  «·—∆Ì”Ì…"
   ClientHeight    =   5730
   ClientLeft      =   75
   ClientTop       =   450
   ClientWidth     =   9480
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   RightToLeft     =   -1  'True
   ScaleHeight     =   5730
   ScaleWidth      =   9480
   Begin VB.OptionButton xInd 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "«· — Ì» ⁄·Ï «·ﬂÊœ"
      Height          =   240
      Index           =   1
      Left            =   2625
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Top             =   900
      Value           =   -1  'True
      Width           =   1740
   End
   Begin VB.OptionButton xInd 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "«· — Ì» ⁄·Ï «·ﬁÌ„…"
      Height          =   240
      Index           =   2
      Left            =   600
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   900
      Width           =   1740
   End
   Begin VB.TextBox xDate2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6825
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   1275
      Width           =   1440
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   825
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      RightToLeft     =   -1  'True
      Top             =   2400
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.TextBox xDate1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6825
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   825
      Width           =   1440
   End
   Begin VB.CommandButton Cmd_Print 
      BackColor       =   &H00C0FFFF&
      Caption         =   "ÿ»«⁄…"
      Enabled         =   0   'False
      Height          =   390
      Left            =   1800
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   75
      Width           =   1515
   End
   Begin VB.CommandButton CmdOk 
      BackColor       =   &H00C0FFFF&
      Caption         =   "⁄—÷"
      Height          =   390
      Left            =   4950
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   75
      Width           =   1515
   End
   Begin VB.CommandButton CmdUndo 
      BackColor       =   &H00C0FFFF&
      Caption         =   " —«Ã⁄"
      Height          =   390
      Left            =   3375
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   75
      Width           =   1515
   End
   Begin VB.CommandButton CmdExit 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Œ—ÊÃ"
      Height          =   390
      Left            =   225
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   75
      Width           =   1515
   End
   Begin Crystal.CrystalReport Report1 
      Left            =   2400
      Top             =   2325
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
   End
   Begin MSFlexGridLib.MSFlexGrid invGrid 
      Height          =   3390
      Left            =   300
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   2025
      Width           =   8940
      _ExtentX        =   15769
      _ExtentY        =   5980
      _Version        =   393216
      Rows            =   1
      Cols            =   1
      FixedCols       =   0
      BackColor       =   12648447
      BackColorFixed  =   32896
      ForeColorFixed  =   16777215
      BackColorSel    =   65535
      BackColorBkg    =   16777215
      Enabled         =   -1  'True
      RightToLeft     =   -1  'True
      ScrollBars      =   2
      BorderStyle     =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "„‰  «—ÌŒ"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   8400
      TabIndex        =   10
      Top             =   900
      Width           =   675
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "≈·Ï  «—ÌŒ"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   8400
      TabIndex        =   9
      Top             =   1350
      Width           =   735
   End
   Begin VB.Label xCode 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3975
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   975
      Visible         =   0   'False
      Width           =   1665
   End
   Begin VB.Label xDesc 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3975
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   1350
      Visible         =   0   'False
      Width           =   1665
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H0000C0C0&
      BorderWidth     =   3
      Height          =   1065
      Left            =   6075
      Shape           =   4  'Rounded Rectangle
      Top             =   675
      Width           =   3240
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0000C0C0&
      BorderWidth     =   3
      Height          =   3690
      Left            =   150
      Top             =   1875
      Width           =   9240
   End
End
Attribute VB_Name = "TSalItem2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim invtable As Recordset
Dim TempTable As Recordset
Dim cString As String
Dim cString2 As String
Dim cFilter1 As String
Dim cFilter As String
Dim nInd As Byte
Private Sub Cmd_Print_Click()
tempdb.Execute "Delete * From Temp"
Set TempTable = tempdb.OpenRecordset("TEMP")
With invtable
If .RecordCount > 0 Then
.MoveFirst
Do While Not .EOF
    TempTable.AddNew
    TempTable.str1 = .CodeGroup
    TempTable.str2 = .Desc_Group
    TempTable.val1 = .SUMQUANT
    TempTable.val2 = .SUMTOTAL
    TempTable.VAL3 = .SUMCOST
    TempTable.STR6 = "»Ì«‰ »≈Ã„«·Ï „»Ì⁄«  «·„Ã„Ê⁄«  «·—∆Ì”Ì… „‰  «—ÌŒ " & xdate1.Text & " ≈·Ï  «—ÌŒ " & xDate2.Text
'   TempTable.str19 = firsttitle
'   TempTable.str20 = Secondtitle
    TempTable.Update
    .MoveNext
Loop
End If
End With
Report1.ReportFileName = PublicPath & "\Reports\TSalItem.rpt"
Report1.DataFiles(0) = App.Path & "\temp.mdb"
Report1.Action = 1
End Sub
Private Sub CmdExit_Click()
Unload Me
Set TSalItem = Nothing
End Sub

Private Sub CmdOk_Click()
If IsDate(xdate1.Text) And IsDate(xDate2.Text) Then
CMD_PRINT.Enabled = True
Me.MousePointer = 11
cFilter = ""
    cString = "SELECT FILE1_70.CODE AS CodeGroup , First(FILE1_70.DESCA) AS Desc_Group,  " & _
              " Sum(FILE6_20.TOTAL) AS SumTOTAL, Sum(FILE6_20.QUANT) AS SumQUANT  , Sum(FILE6_20.QUANT * FILE6_20.COST ) AS SumCost " & _
              " FROM FILE1_70 RIGHT JOIN ((FILE1_50 RIGHT JOIN FILE1_10 ON FILE1_50.CODE = FILE1_10.GROUP) RIGHT JOIN FILE6_20 ON FILE1_10.ITEM = FILE6_20.ITEM) ON FILE1_70.CODE = FILE1_50.M_GROUP " & _
              " WHERE FILE1_70.FLAG=2 " & _
              " AND [Date] >= DateValue('" & xdate1.Text & "')" & _
              " AND [DatE] <= DateValue('" & xDate2.Text & "')" & _
              " GROUP BY FILE1_70.CODE "

If nInd = 1 Then
    cString = cString & " ORDER BY FILE1_70.CODE  "
Else
    cString = cString & " ORDER BY Sum(FILE6_20.TOTAL)  DESC "
End If
Set invtable = mydb.CreateSnapshot(cString)
If invtable.RecordCount > 0 Then
    Fillgrd
Else
    invGrid.Rows = 1
End If
Me.MousePointer = 0
Else
 MsgBox ("—«Ã⁄  ”ÃÌ· «· «—ÌŒ À„ ⁄«Êœ «·„Õ«Ê·…")
End If
End Sub
Sub Fillgrd()
I = 1
invGrid.Rows = 1
If formMode <> EmptyMode Then invGrid.Rows = 1
invtable.MoveFirst
Do
   invGrid.AddItem ""
   invGrid.TextArray(faIndex(I, 0, invGrid)) = TurnValue(invtable.CodeGroup, Null, "")
   invGrid.TextArray(faIndex(I, 1, invGrid)) = TurnValue(invtable.Desc_Group, Null, "")
   invGrid.TextArray(faIndex(I, 2, invGrid)) = TurnValue(invtable.SUMQUANT, Null, "")
   invGrid.TextArray(faIndex(I, 3, invGrid)) = TurnValue(Format(invtable.SUMTOTAL, "##0.00"), Null, "")
   invGrid.TextArray(faIndex(I, 4, invGrid)) = TurnValue(Format(invtable.SUMTOTAL - invtable.SUMCOST, "##0.00"), Null, "")
   invtable.MoveNext
   I = I + 1
Loop Until invtable.EOF
End Sub
Private Sub CmdUndo_Click()
    Unload Me
End Sub
Private Sub Form_Load()
xdate1.Text = ""
xDate2.Text = ""
'Me.Picture = LoadPicture(App.Path & "\mainback.jpg")
invGrid.FormatString = "ﬂÊœ|" & "«·»Ì«‰|" & "≈Ã„«·Ï ﬂ„Ì…|" & " ﬁÌ„… „»Ì⁄« |" & "≈Ã„«·Ï —»Õ"
invGrid.Cols = 5
invGrid.ColWidth(0) = 1000
invGrid.ColWidth(1) = 3500
invGrid.ColWidth(2) = 1500
invGrid.ColWidth(3) = 1500
invGrid.ColWidth(4) = 1500

For I = 0 To invGrid.Cols - 1
    invGrid.ColAlignment(I) = 1
Next
nInd = 1
End Sub
Private Sub InvGrid_Click()
    If invGrid.Rows > 1 Then
        xCode.Caption = GrdText(invGrid, 0)
        xDesc.Caption = GrdText(invGrid, 1)
        Load TGrItem
        TGrItem.Show 1
    End If
End Sub
Private Sub xInd_Click(Index As Integer)
If xInd(Index).Value = True Then nInd = Index
End Sub
