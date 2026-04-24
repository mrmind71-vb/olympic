VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form aDDdAY 
   BackColor       =   &H00E0E0E0&
   Caption         =   " ”ÃÌ· ﬂ‘› «·Õ÷Ê— Ê «·≈‰’—«›"
   ClientHeight    =   6750
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9435
   FillStyle       =   4  'Upward Diagonal
   LinkTopic       =   "Form1"
   RightToLeft     =   -1  'True
   ScaleHeight     =   6750
   ScaleWidth      =   9435
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox xCode 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1980
      RightToLeft     =   -1  'True
      TabIndex        =   19
      Top             =   150
      Visible         =   0   'False
      Width           =   630
   End
   Begin VB.TextBox xAddSal 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   330
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   1320
      Width           =   630
   End
   Begin VB.TextBox xFriday 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1485
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   1320
      Width           =   630
   End
   Begin VB.TextBox xDebit 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2640
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   1320
      Width           =   630
   End
   Begin VB.TextBox xDisc 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   3840
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   1320
      Width           =   630
   End
   Begin VB.CommandButton Cmd_Exit 
      BackColor       =   &H00C0C000&
      Cancel          =   -1  'True
      Caption         =   "Œ—ÊÃ"
      CausesValidation=   0   'False
      Height          =   390
      Left            =   480
      MaskColor       =   &H00FFFFFF&
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   75
      Width           =   1140
   End
   Begin VB.CommandButton Cmd_Save 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Õ›Ÿ"
      Height          =   390
      Left            =   240
      MaskColor       =   &H00000080&
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1800
      Width           =   1140
   End
   Begin VB.TextBox xTime 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   4965
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   1320
      Width           =   630
   End
   Begin VB.TextBox xName 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   315
      Left            =   5940
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   1305
      Width           =   3105
   End
   Begin VB.CommandButton Cmd_Ok 
      BackColor       =   &H00C0C000&
      Caption         =   " ”ÃÌ·"
      Height          =   315
      Left            =   4440
      MaskColor       =   &H000000FF&
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   240
      Width           =   1140
   End
   Begin VB.TextBox xDate 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   7050
      TabIndex        =   0
      Top             =   225
      Width           =   1320
   End
   Begin MSFlexGridLib.MSFlexGrid InvGrid 
      Height          =   4410
      Left            =   120
      TabIndex        =   16
      Top             =   2250
      Width           =   9225
      _ExtentX        =   16272
      _ExtentY        =   7779
      _Version        =   393216
      BackColor       =   14737632
      BackColorFixed  =   12632064
      BackColorSel    =   12648447
      BackColorBkg    =   14737632
      RightToLeft     =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00000040&
      X1              =   9120
      X2              =   0
      Y1              =   1725
      Y2              =   1725
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00000040&
      X1              =   9120
      X2              =   300
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000040&
      X1              =   9120
      X2              =   60
      Y1              =   675
      Y2              =   675
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ÕÊ«›“"
      Height          =   195
      Left            =   465
      TabIndex        =   18
      Top             =   840
      Width           =   435
   End
   Begin VB.Line Line10 
      BorderColor     =   &H00000040&
      X1              =   1215
      X2              =   1215
      Y1              =   1725
      Y2              =   675
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ã„⁄…"
      Height          =   195
      Left            =   1560
      TabIndex        =   17
      Top             =   840
      Width           =   375
   End
   Begin VB.Line Line9 
      BorderColor     =   &H00000040&
      X1              =   2370
      X2              =   2370
      Y1              =   1725
      Y2              =   675
   End
   Begin VB.Label xDay 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   6000
      RightToLeft     =   -1  'True
      TabIndex        =   15
      Top             =   225
      Width           =   915
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "”·›Ì…"
      Height          =   195
      Left            =   2760
      TabIndex        =   14
      Top             =   840
      Width           =   405
   End
   Begin VB.Line Line8 
      BorderColor     =   &H00000040&
      X1              =   3525
      X2              =   3525
      Y1              =   1725
      Y2              =   675
   End
   Begin VB.Line Line7 
      BorderColor     =   &H00000040&
      X1              =   60
      X2              =   60
      Y1              =   1725
      Y2              =   675
   End
   Begin VB.Line Line6 
      BorderColor     =   &H00000040&
      X1              =   4695
      X2              =   4695
      Y1              =   1725
      Y2              =   675
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00000040&
      X1              =   5850
      X2              =   5850
      Y1              =   1725
      Y2              =   675
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "«·Õ÷Ê—"
      Height          =   195
      Left            =   4815
      TabIndex        =   12
      Top             =   840
      Width           =   525
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Œ’„"
      Height          =   195
      Left            =   3840
      TabIndex        =   11
      Top             =   840
      Width           =   345
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "«·≈”„"
      Height          =   195
      Left            =   7425
      TabIndex        =   10
      Top             =   840
      Width           =   375
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00000040&
      X1              =   9120
      X2              =   9120
      Y1              =   1725
      Y2              =   675
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "«· «—ÌŒ"
      Height          =   195
      Left            =   8460
      TabIndex        =   9
      Top             =   270
      Width           =   495
   End
End
Attribute VB_Name = "aDDdAY"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim EmpTable As Recordset
Dim PresTable As Recordset
Dim DayTable As Recordset
Dim cString As String
Private Sub Cmd_Exit_Click()
Unload Me
End Sub
Private Sub Cmd_OK_Click()
Fillgrd
EmpTable.MoveFirst
xName.Text = EmpTable.NAMEA
xCode.Text = EmpTable.CODE
PresTable.Index = "FILE9_20_1"
PresTable.Seek "=", DateValue(xDate.Text), xCode.Text
If Not PresTable.NoMatch Then
    xTime.Text = TurnValue(PresTable.Time, Null, "")
    xDisc.Text = TurnValue(PresTable.DISC, Null, "")
    xFriday.Text = TurnValue(PresTable.Friday, Null, "")
    xDebit.Text = TurnValue(PresTable.DEBIT, Null, "")
    xAddSal.Text = TurnValue(PresTable.ADDSAL, Null, "")
Else
    xTime.Text = ""
    xDisc.Text = ""
    xFriday.Text = ""
    xDebit.Text = ""
    xAddSal.Text = ""
End If
xTime.SetFocus
End Sub
Private Sub Cmd_Save_Click()
PresTable.Seek "=", DateValue(xDate.Text), xCode.Text
If PresTable.NoMatch Then
    PresTable.AddNew
Else
    PresTable.Edit
End If
PresTable.CODE = xCode.Text
PresTable.NAMEA = EmpTable.NAMEA
PresTable.Date = TurnValue(DateValue(xDate.Text), "", Null)
PresTable.Time = TurnValue(Val(xTime.Text), "", Null)
PresTable.DISC = TurnValue(Val(xDisc.Text), "", Null)
PresTable.Friday = TurnValue(Val(xFriday.Text), "", Null)
PresTable.DEBIT = TurnValue(Val(xDebit.Text), "", Null)
PresTable.ADDSAL = TurnValue(Val(xAddSal.Text), "", Null)
PresTable.Update
EmpTable.MoveNext
If EmpTable.EOF Or Not IsNull(EmpTable.enddate) Then
     Unload Me
Else
'   If EmpTable.Fix = 1 Then
    PresTable.Seek "=", DateValue(xDate.Text), EmpTable.CODE
    If Not PresTable.NoMatch Then
        xName.Text = EmpTable.NAMEA
        xCode.Text = EmpTable.CODE
        xTime.Text = TurnValue(PresTable.Time, Null, "")
        xDisc.Text = TurnValue(PresTable.DISC, Null, "")
        xFriday.Text = TurnValue(PresTable.Friday, Null, "")
        xDebit.Text = TurnValue(PresTable.DEBIT, Null, "")
        xAddSal.Text = TurnValue(PresTable.ADDSAL, Null, "")
        xTime.SetFocus
    Else
        xName.Text = EmpTable.NAMEA
        xCode.Text = EmpTable.CODE
'       xTime.Text = ""
        xDisc.Text = ""
        xFriday.Text = ""
        xAddSal.Text = ""
        xDebit.Text = ""
        xTime.SetFocus
    End If
'   End If
End If
Fillgrd
End Sub
Private Sub Form_Load()
'Me.Picture = LoadPicture(App.Path & "\GRAPH\MAIN.jpg")
Set EmpTable = mydb.OpenRecordset("file9_10")
Set PresTable = mydb.OpenRecordset("file9_20")

xDay.Caption = ""
xDate.Text = Format(Date, "DD-MM-YYYY")
xDay.Caption = NameOfDay(xDate.Text)
xTime.Text = ""
xDisc.Text = ""
xName.Text = ""
xCode.Text = ""
xFriday.Text = ""
xAddSal.Text = ""
InvGrid.FormatString = "≈”„ «·„ÊŸ›|" & "«·Õ÷Ê—|" & "«·Œ’„|" & "”·›Ì…|" & "Ã„⁄…|" & "ÕÊ«›“|" & "ﬂÊœ"
InvGrid.Cols = 7
InvGrid.ColWidth(0) = 3500
InvGrid.ColWidth(1) = 1000
InvGrid.ColWidth(2) = 1000
InvGrid.ColWidth(3) = 1000
InvGrid.ColWidth(4) = 1000
InvGrid.ColWidth(5) = 1000
InvGrid.ColWidth(6) = 0
End Sub
Private Sub invGrid_DblClick()
    xName.Text = GrdText(InvGrid, 0)
    xTime.Text = GrdText(InvGrid, 1)
    xDisc.Text = GrdText(InvGrid, 2)
    xDebit.Text = GrdText(InvGrid, 3)
    xFriday.Text = GrdText(InvGrid, 4)
    xAddSal.Text = GrdText(InvGrid, 5)
    xCode.Text = GrdText(InvGrid, 6)
    xTime.SetFocus
End Sub
Private Sub xDate_Validate(Cancel As Boolean)
If IsDate(xDate.Text) Then
    xDay.Caption = NameOfDay(xDate.Text)
Else
    Cancel = True
End If
End Sub
Sub Fillgrd()
'cString = " SELECT FILE_5.CODE, FILE_5.NAMEA, file_7.date, file_7.late, file_7.type, file_7.going FROM FILE_5 LEFT JOIN file_7 ON FILE_5.CODE = file_7.CODE WHERE FILE_7.[DATE] = " & MyParn(DateValue(xDate.Text))
cString = " SELECT FILE9_10.CODE, FILE9_10.NAMEA, FILE9_20.date, FILE9_20.TIME, FILE9_20.DEBIT, FILE9_20.DISC , FILE9_20.Friday , FILE9_20.ADDSAL FROM FILE9_10 LEFT JOIN FILE9_20 ON FILE9_10.CODE = FILE9_20.CODE " & _
                    " where Date = DateValue(" & MyParn(xDate.Text) & ")"
Set DayTable = mydb.OpenRecordset(cString, dbOpenDynaset)
InvGrid.Rows = 1
I = 1
If DayTable.RecordCount = 0 Then Exit Sub
DayTable.MoveFirst
If DayTable.NoMatch Then Exit Sub
Do While True
    InvGrid.AddItem ""
    InvGrid.TextArray(faIndex(I, 0, InvGrid)) = TurnValue(DayTable.NAMEA, Null, "")
    InvGrid.TextArray(faIndex(I, 1, InvGrid)) = TurnValue(DayTable.Time, Null, "")
    InvGrid.TextArray(faIndex(I, 2, InvGrid)) = TurnValue(DayTable.DISC, Null, "")
    InvGrid.TextArray(faIndex(I, 3, InvGrid)) = TurnValue(DayTable.DEBIT, Null, "")
    InvGrid.TextArray(faIndex(I, 4, InvGrid)) = TurnValue(DayTable.Friday, Null, "")
    InvGrid.TextArray(faIndex(I, 5, InvGrid)) = TurnValue(DayTable.ADDSAL, Null, "")
    InvGrid.TextArray(faIndex(I, 6, InvGrid)) = TurnValue(DayTable.CODE, Null, "")
    DayTable.MoveNext
    I = I + 1
If DayTable.EOF Then Exit Do
Loop
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is DBCombo Then SendKeys "{TAB}"
End If
End Sub
Private Sub invGrid_KeyUp(KeyCode As Integer, Shift As Integer)
If InvGrid.Rows <= 2 Then Exit Sub
If KeyCode = 46 Then
    If MsgBox("«·€«¡ «·⁄«„· !! Â· «‰  „Ê«›ﬁ ø", 1 + 256) = vbOK Then deleteLine
    Exit Sub
End If
If KeyCode = 13 Then invGrid_DblClick
End Sub
Sub deleteLine()
PresTable.Seek "=", DateValue(xDate.Text), GrdText(InvGrid, 6)
If Not PresTable.NoMatch Then
    PresTable.Delete
End If
Fillgrd
End Sub

