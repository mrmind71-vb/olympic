VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form salesmanfrm 
   Caption         =   "«”„«¡ «·»«∆⁄Ì‰"
   ClientHeight    =   495
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4980
   LinkTopic       =   "Form1"
   ScaleHeight     =   495
   ScaleWidth      =   4980
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdapply 
      Caption         =   "œŒÊ·"
      Height          =   330
      Left            =   45
      TabIndex        =   2
      Top             =   90
      Width           =   1365
   End
   Begin MSDataListLib.DataCombo xman 
      Height          =   315
      Left            =   1485
      TabIndex        =   0
      Top             =   90
      Width           =   2760
      _ExtentX        =   4868
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSAdodcLib.Adodc DATA1 
      Height          =   330
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   2340
      _ExtentX        =   4128
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
      Caption         =   "data1"
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
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "«·»«∆⁄ :"
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
      Left            =   4365
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   135
      Width           =   495
   End
End
Attribute VB_Name = "salesmanfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdApply_Click()
If Trim(xman.BoundText) = "" Then
    MsgBox "·«»œ „‰  ÕœÌœ «·»«∆⁄"
Else
    cSalesMan = xman.BoundText
    Unload Me
    Main.Show 1
End If
End Sub
Private Sub cmdapply_LostFocus()
If Not xman.MatchedWithList Then xman.BoundText = "'"
End Sub
Private Sub Form_Load()
DATA1.ConnectionString = CON.ConnectionString
DATA1.RecordSource = "FILE6_25"
Set xman.RowSource = DATA1
xman.ListField = "Desca"
xman.BoundColumn = "Code"
End Sub

Private Sub Form_Terminate()
End
End Sub
