VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form advancefrm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "»Ì«‰ «·”·›Ì« "
   ClientHeight    =   4095
   ClientLeft      =   405
   ClientTop       =   1455
   ClientWidth     =   7350
   FillColor       =   &H00808080&
   FillStyle       =   0  'Solid
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
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   RightToLeft     =   -1  'True
   ScaleHeight     =   4095
   ScaleWidth      =   7350
   Begin VB.Frame Frame8 
      Height          =   600
      Left            =   180
      RightToLeft     =   -1  'True
      TabIndex        =   19
      Top             =   3330
      Width           =   3255
      Begin Threed.SSCommand cmdLast 
         CausesValidation=   0   'False
         Height          =   420
         Left            =   45
         TabIndex        =   20
         Top             =   135
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   741
         _Version        =   196610
         PictureFrames   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "advance.frx":0000
         Caption         =   "«ŒÌ—"
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "advance.frx":21D0
      End
      Begin Threed.SSCommand cmdNext 
         CausesValidation=   0   'False
         Height          =   420
         Left            =   825
         TabIndex        =   21
         Top             =   135
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   741
         _Version        =   196610
         PictureFrames   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "advance.frx":4318
         Caption         =   "·«Õﬁ "
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "advance.frx":64E0
      End
      Begin Threed.SSCommand cmdPrevious 
         CausesValidation=   0   'False
         Height          =   420
         Left            =   1605
         TabIndex        =   22
         Top             =   135
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   741
         _Version        =   196610
         PictureFrames   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "advance.frx":862F
         Caption         =   "”«»ﬁ"
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "advance.frx":A80F
      End
      Begin Threed.SSCommand cmdFirst 
         CausesValidation=   0   'False
         Height          =   420
         Left            =   2385
         TabIndex        =   23
         Top             =   135
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   741
         _Version        =   196610
         PictureFrames   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "advance.frx":C96A
         Caption         =   "√Ê·"
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "advance.frx":EB26
      End
   End
   Begin VB.Frame Frame4 
      Height          =   690
      Left            =   45
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Top             =   -45
      Width           =   7215
      Begin VB.CommandButton CmdInform 
         CausesValidation=   0   'False
         Height          =   510
         Left            =   5985
         Picture         =   "advance.frx":10C75
         Style           =   1  'Graphical
         TabIndex        =   18
         TabStop         =   0   'False
         ToolTipText     =   "«” ⁄·«„"
         Top             =   135
         Width           =   1185
      End
      Begin VB.CommandButton cmdAdd 
         CausesValidation=   0   'False
         Height          =   510
         Left            =   4785
         MaskColor       =   &H00FFFFFF&
         Picture         =   "advance.frx":13448
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   17
         TabStop         =   0   'False
         ToolTipText     =   "«÷«›…"
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1185
      End
      Begin VB.CommandButton CmdDel 
         CausesValidation=   0   'False
         Height          =   510
         Left            =   1230
         MaskColor       =   &H00FFFFFF&
         Picture         =   "advance.frx":159F4
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   16
         TabStop         =   0   'False
         ToolTipText     =   "Õ–›"
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1185
      End
      Begin VB.CommandButton CmdExit 
         CausesValidation=   0   'False
         Height          =   510
         Left            =   45
         MaskColor       =   &H00FFFFFF&
         Picture         =   "advance.frx":1828E
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   15
         TabStop         =   0   'False
         ToolTipText     =   "Œ—ÊÃ"
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1185
      End
      Begin VB.CommandButton CmdUndo 
         CausesValidation=   0   'False
         Height          =   510
         Left            =   2415
         MaskColor       =   &H00FFFFFF&
         Picture         =   "advance.frx":1A6FA
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   14
         TabStop         =   0   'False
         ToolTipText     =   " —«Ã⁄"
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1185
      End
      Begin VB.CommandButton cmdSave 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   3600
         MaskColor       =   &H00FFFFFF&
         Picture         =   "advance.frx":1CC73
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   13
         TabStop         =   0   'False
         ToolTipText     =   "Õ›Ÿ"
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1185
      End
   End
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   585
      Top             =   630
      Visible         =   0   'False
      Width           =   1740
      _ExtentX        =   3069
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.TextBox xCode 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "Arabic Transparent"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   5040
      Locked          =   -1  'True
      MaxLength       =   6
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   765
      Width           =   1365
   End
   Begin VB.TextBox xDate 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "Arabic Transparent"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   5040
      MaxLength       =   10
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   1170
      Width           =   1365
   End
   Begin VB.TextBox xValue 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "Arabic Transparent"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   4860
      MaxLength       =   10
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   2700
      Width           =   1545
   End
   Begin VB.TextBox xDescA 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "Arabic Transparent"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   225
      MaxLength       =   100
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   2295
      Width           =   6180
   End
   Begin MSDataListLib.DataCombo xBox 
      Height          =   330
      Left            =   2070
      TabIndex        =   2
      Top             =   1575
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   582
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      Text            =   "DataCombo1"
      RightToLeft     =   -1  'True
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
   Begin MSDataListLib.DataCombo xDriver 
      Height          =   330
      Left            =   2070
      TabIndex        =   3
      Top             =   1935
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   582
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      Text            =   "DataCombo1"
      RightToLeft     =   -1  'True
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
   Begin MSAdodcLib.Adodc DATA2 
      Height          =   330
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1740
      _ExtentX        =   3069
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H80000013&
      BackStyle       =   0  'Transparent
      Caption         =   "„”·”·"
      BeginProperty Font 
         Name            =   "Arabic Transparent"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   6525
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   765
      Width           =   525
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H80000013&
      BackStyle       =   0  'Transparent
      Caption         =   "«·„ÊŸ›"
      BeginProperty Font 
         Name            =   "Arabic Transparent"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   1
      Left            =   6525
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   1980
      Width           =   600
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000013&
      BackStyle       =   0  'Transparent
      Caption         =   "«·ﬁÌ„…"
      BeginProperty Font 
         Name            =   "Arabic Transparent"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   6525
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   2835
      Width           =   435
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H80000013&
      BackStyle       =   0  'Transparent
      Caption         =   " «—ÌŒ"
      BeginProperty Font 
         Name            =   "Arabic Transparent"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   6525
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   1170
      Width           =   390
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H80000013&
      BackStyle       =   0  'Transparent
      Caption         =   "»Ì«‰"
      BeginProperty Font 
         Name            =   "Arabic Transparent"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   6525
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   2385
      Width           =   300
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackColor       =   &H80000013&
      BackStyle       =   0  'Transparent
      Caption         =   "„‰ Œ“‰…"
      BeginProperty Font 
         Name            =   "Arabic Transparent"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   6525
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   1620
      Width           =   630
   End
End
Attribute VB_Name = "advancefrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public bEdit As Boolean
Dim formMode As Byte
Dim con As New ADODB.Connection
Dim oSearch As New Search3, oSearchBox As New Search3, oSearchDriver As New Search3
Dim CardTable As ADODB.Recordset
Const LoadMode = 1, DefineMode = 2
Sub Handlecontrols(nMode)
cmdAdd.Enabled = (nMode = LoadMode And bEdit)
CmdDel.Enabled = (nMode = LoadMode And bEdit)
cmdSave.Enabled = bEdit
CmdInform.Enabled = (nMode = LoadMode)
cmdPrevious.Enabled = (nMode = LoadMode) And CardTable.AbsolutePosition > 1
cmdNext.Enabled = (nMode = LoadMode) And CardTable.AbsolutePosition < CardTable.RecordCount
cmdLast.Enabled = (nMode = LoadMode) And CardTable.AbsolutePosition < CardTable.RecordCount And CardTable.RecordCount > 2
cmdFirst.Enabled = (nMode = LoadMode) And CardTable.AbsolutePosition > 1 And CardTable.RecordCount > 2
XCODE.Enabled = Not (nMode = LoadMode)
End Sub
Private Sub CardLookup()
Dim Generalarray(5)
Dim listarray(3, 5)
Dim GrdArray(5, 1)
Set Generalarray(0) = Me
Dim cString As String
cString = "SELECT FILE2_52.CODE,FILE0_50.DESCA,DRIVER.DESCA,FILE2_52.DESCA,CONVERT(VARCHAR(10),[DATE],111),[VALUE]" & _
          " FROM FILE2_52 INNER JOIN FILE0_50 ON FILE2_52.BOX = FILE0_50.CODE INNER JOIN DRIVER ON FILE2_52.CODE = DRIVER.CODE"

Generalarray(1) = cString
Generalarray(2) = "ORDER BY FILE2_52.DATE"
Generalarray(3) = 5000
Generalarray(5) = False

listarray(0, 0) = "«·»Ì«‰ √Ê «· «—ÌŒ"
listarray(0, 1) = "(%%FILE2_52.DESCA%% or ##date##)"

listarray(1, 0) = "„‰ Œ“‰…"
listarray(1, 1) = "(%%FILE0_50.DESCA%%)"

listarray(2, 0) = "«·„ÊŸ›"
listarray(2, 1) = "(%%DRIVER.DESCA%%)"

listarray(3, 0) = "«·ﬁÌ„…"
listarray(3, 1) = "(**[VALUE]**)"

GrdArray(0, 0) = "«·ﬂÊœ"
GrdArray(0, 1) = 0

GrdArray(1, 0) = "„‰ Œ“‰…"
GrdArray(1, 1) = 2000

GrdArray(2, 0) = "«·„ÊŸ›"
GrdArray(2, 1) = 2000

GrdArray(3, 0) = "«·»Ì«‰"
GrdArray(3, 1) = 3000

GrdArray(4, 0) = "«· «—ÌŒ"
GrdArray(4, 1) = 1400

GrdArray(5, 0) = "«·ﬁÌ„…"
GrdArray(5, 1) = 1000

searchArray = Array(Generalarray, listarray, GrdArray)
oSearch.Caption = "≈” ⁄·«„ «·Œ“‰"
oSearch.Show 1
End Sub
Sub mydefine()
XCODE.Text = RetZero(Val(Newflag("FILE2_52", "CODE")), 6)
xDesca.Text = ""
xDate.Text = ""
XBOX.BoundText = ""
xDriver.BoundText = ""
xValue.Text = ""
Handlecontrols DefineMode
End Sub
Sub myProc()
If ActiveControl.Name = CmdInform.Name Then
    XCODE.Text = oSearch.grid1.TextMatrix(oSearch.grid1.Row, 0)
    Unload oSearch
    myUndo
ElseIf ActiveControl.Name = XBOX.Name Then
    XBOX.BoundText = oSearchBox.grid1.TextMatrix(oSearchBox.grid1.Row, 0)
    Unload oSearchBox
ElseIf ActiveControl.Name = xDriver.Name Then
    xDriver.BoundText = oSearchDriver.grid1.TextMatrix(oSearchDriver.grid1.Row, 0)
    Unload oSearchDriver
End If
End Sub
Sub myload()
XCODE.Text = CardTable!Code
xDesca.Text = CardTable!Desca & ""
XBOX.BoundText = CardTable!BOX & ""
xDriver.BoundText = CardTable!Driver & ""
xDate.Text = Format(CardTable!Date, "YYYY-MM-DD")
xValue.Text = Myvalue(CardTable!Value, "Fixed")
Handlecontrols LoadMode
End Sub
Private Function myreplace() As Boolean
Dim aInsert As Variant
aInsert = AddFlag(Empty, "DESCA", addstring(xDesca.Text))
aInsert = AddFlag(aInsert, "[DATE]", addDate(xDate.Text))
aInsert = AddFlag(aInsert, "[BOX]", addstring(XBOX.BoundText))
aInsert = AddFlag(aInsert, "[DRIVER]", addstring(xDriver.BoundText))
aInsert = AddFlag(aInsert, "[VALUE]", Val(xValue.Text))
con.BeginTrans
On Error GoTo myerror
If XCODE.Enabled Then
    XCODE.Text = RetZero(Val(Newflag("FILE2_52", "CODE")), 6)
    aInsert = AddFlag(aInsert, "CODE", addstring(XCODE.Text))
    con.Execute addInsert(aInsert, "FILE2_52")
Else
    con.Execute addUpdate(aInsert, "FILE2_52", "CODE = " & addstring(XCODE.Text))
End If
con.CommitTrans
myreplace = True
Exit Function
myerror:
MsgBox Err.Description
con.RollbackTrans
Err.Clear
End Function
Function MYVALID() As Boolean
If XCODE.Text = "" Then
    MsgBox " ”ÃÌ· „”·”· "
    Exit Function
End If
If xDate.Text = "" Then
    MsgBox " ”ÃÌ·  «—ÌŒ "
    Exit Function
End If
If Trim(XBOX.BoundText) = "" Then
    MsgBox " ”ÃÌ· «·Œ“«‰… «·«Ê·Ì ÷—Ê—Ì"
    Exit Function
End If
If Trim(xDriver.BoundText) = "" Then
    MsgBox " ”ÃÌ· «·Œ“«‰… «·À«‰Ì… ÷—Ê—Ì"
    Exit Function
End If
MYVALID = True
End Function
Private Sub CmdAdd_Click()
mydefine
On Error Resume Next
XCODE.SetFocus
Err.Clear
End Sub
Private Sub CmdDel_Click()
If MsgBox("«·€«¡ «·”Ã· «·Õ«·Ï : Â· «‰  „Ê«›ﬁ ø", 4) = 6 Then
    On Error GoTo myerror
    con.BeginTrans
    con.Execute "Delete  From FILE2_52 Where Code = " & MyParn(XCODE.Text)
    con.CommitTrans
    openCardTable
    If Not (CardTable.EOF And CardTable.BOF) Then
        CardTable.Find "Code < " & MyParn(XCODE.Text), , adSearchBackward, adBookmarkLast
        If CardTable.BOF Then CardTable.MoveFirst
        myload
    Else
        mydefine
    End If
End If
Exit Sub
myerror:
con.RollbackTrans
MsgBox Err.Description
Err.Clear
End Sub
Private Sub cmdExit_Click()
    Unload Me
End Sub
Private Sub CmdFirst_Click()
CardTable.MoveFirst
myload
End Sub
Private Sub CmdInform_Click()
CardLookup
End Sub
Private Sub CmdLast_Click()
CardTable.MoveLast
myload
End Sub
Private Sub CmdNext_Click()
CardTable.MoveNext
If CardTable.EOF Then
    CardTable.MovePrevious
Else
    myload
End If
End Sub
Private Sub CmdPrevious_Click()
CardTable.MovePrevious
If CardTable.BOF Then
    CardTable.MoveNext
Else
    myload
End If
End Sub
Private Sub cmdSave_Click()
If Not MYVALID Then Exit Sub
If Not myreplace Then Exit Sub
Inform " „ Õ›Ÿ «·»Ì«‰«  »‰Ã«Õ"
If XCODE.Enabled Then
    CmdAdd_Click
Else
    openCardTable
    myUndo
End If
End Sub
Private Sub CmdUndo_Click()
openCardTable
myUndo
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is DataCombo Then SendKeys "{TAB}"
ElseIf KeyAscii = 19 And cmdSave.Enabled Then
    cmdSave_Click
End If
End Sub
Private Sub Form_Load()
bEdit = True
openCon con
data1.ConnectionString = strCon
data1.RecordSource = "Select * From file0_50 where code > '500000'"

DATA2.ConnectionString = strCon
DATA2.RecordSource = "Select * From driver"

Set XBOX.RowSource = data1
XBOX.ListField = "Desca"
XBOX.BoundColumn = "Code"

Set xDriver.RowSource = DATA2
xDriver.ListField = "Desca"
xDriver.BoundColumn = "Code"

openCardTable
myUndo
End Sub
Private Sub Form_Unload(Cancel As Integer)
If MsgBox(cMsgExit, vbOKCancel + vbDefaultButton1) = vbOK Then
    cmdSave_Click
End If
CardTable.Close
Set CardTable = Nothing
closeCon con
End Sub

Private Sub xBox_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
    BoxLookupAll Me, oSearchBox, "FILE0_50.CODE < '500001'"
End If
End Sub

Private Sub xDriver_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
    DriverLookupAll Me, oSearchDriver
End If

End Sub

Private Sub xCode_LostFocus()
XCODE.Text = RetZero(XCODE.Text, 6)
CardTable.Find "Code = " & MyParn(XCODE.Text), , adSearchForward, adBookmarkFirst
If Not CardTable.EOF Then myload
myLostFocus XCODE
End Sub
Private Sub xCode_GotFocus()
myGotFocus XCODE
End Sub
Private Sub xDate_GotFocus()
myGotFocus xDate
End Sub
Private Sub xDate_LostFocus()
myLostFocus xDate
End Sub
Private Sub xDate_Validate(Cancel As Boolean)
myValidDate xDate
End Sub

Private Sub xDesca_LostFocus()
myLostFocus xDesca
End Sub

Private Sub xValue_GotFocus()
myGotFocus xValue
End Sub
Private Sub xDescA_GotFocus()
myGotFocus xDesca
End Sub
Private Sub myUndo()
If (CardTable.BOF And CardTable.EOF) Then
    mydefine
Else
    If Trim(XCODE.Text) <> "" Then
        CardTable.Find "CODE = " & MyParn(XCODE.Text), , adSearchForward, adBookmarkFirst
        If CardTable.EOF Then CardTable.MoveLast
    Else
        CardTable.MoveLast
    End If
    myload
End If
End Sub
Private Sub openCardTable()
Dim cString As String
cString = "SELECT * FROM FILE2_52"
cString = cString & " order by code"
Set CardTable = New ADODB.Recordset
CardTable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
End Sub
Private Sub xValue_LostFocus()
myLostFocus xValue
End Sub
