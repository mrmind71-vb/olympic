VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form Inv_Move 
   BackColor       =   &H00E0E0E0&
   Caption         =   "≈Ã„«·Ï „»Ì⁄«  ⁄„Ì· "
   ClientHeight    =   5760
   ClientLeft      =   75
   ClientTop       =   450
   ClientWidth     =   9390
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   RightToLeft     =   -1  'True
   ScaleHeight     =   5760
   ScaleWidth      =   9390
   Begin Crystal.CrystalReport Report1 
      Left            =   2160
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2880
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      RightToLeft     =   -1  'True
      Top             =   120
      Visible         =   0   'False
      Width           =   1140
   End
   Begin MSFlexGridLib.MSFlexGrid invGrid 
      Height          =   4140
      Left            =   300
      TabIndex        =   5
      Top             =   1275
      Width           =   8790
      _ExtentX        =   15505
      _ExtentY        =   7303
      _Version        =   393216
      Rows            =   1
      Cols            =   1
      FixedCols       =   0
      BackColor       =   12648447
      BackColorFixed  =   32896
      ForeColorFixed  =   16777215
      BackColorSel    =   65535
      BackColorBkg    =   14737632
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
   Begin VB.Shape Shape1 
      BorderColor     =   &H0000C0C0&
      BorderWidth     =   3
      Height          =   4515
      Left            =   150
      Top             =   1125
      Width           =   9165
   End
   Begin VB.Label xType 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   4725
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   75
      Visible         =   0   'False
      Width           =   2040
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   " «—ÌŒ «·›« Ê—…"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   8385
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   600
      Width           =   840
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "—ﬁ„ «·›« Ê—…"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   8340
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   75
      Width           =   780
   End
   Begin VB.Label xDate 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   6900
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   525
      Width           =   1290
   End
   Begin VB.Label xInv 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   6600
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   75
      Width           =   1290
   End
End
Attribute VB_Name = "Inv_Move"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim InvTable As Recordset
Dim TempTable As Recordset
Dim cString As String
Sub Fillgrd()
I = 1
InvGrid.Rows = 1
If formMode <> EmptyMode Then InvGrid.Rows = 1
InvTable.MoveFirst
Do
   InvGrid.AddItem ""
   InvGrid.TextArray(faIndex(I, 0, InvGrid)) = TurnValue(InvTable.Store, Null, "")
   InvGrid.TextArray(faIndex(I, 1, InvGrid)) = TurnValue(InvTable.Item, Null, "")
   InvGrid.TextArray(faIndex(I, 2, InvGrid)) = TurnValue(InvTable.ItemDesc, Null, "")
   InvGrid.TextArray(faIndex(I, 3, InvGrid)) = TurnValue(InvTable.Quant, Null, "")
   InvGrid.TextArray(faIndex(I, 4, InvGrid)) = TurnValue(InvTable.price, Null, "")
   InvGrid.TextArray(faIndex(I, 5, InvGrid)) = TurnValue(InvTable.total, Null, "")
   InvTable.MoveNext
   I = I + 1
Loop Until InvTable.EOF
End Sub
Private Sub CmdUndo_Click()
    Unload Me
End Sub

Private Sub Form_Load()
Dim TargetTable As Recordset
xInv.Caption = ClientMove.xInv.Caption
'xDate.Caption = ClientMove.xDate.Caption
xType.Caption = ClientMove.xType.Caption

tempdb.Execute "DELETE * FROM TEMP"
Set TargetTable = tempdb.CreateDynaset("TEMP")
If publicFlag = 1 Then
    Select Case xType.Caption
        Case 4
            Me.Caption = " „»Ì‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹⁄«  "
            cString = "SELECT FILE6_20.STORE AS STORE , FILE6_20.TOTAL AS TOTAL , FILE6_20.PRICE AS PRICE , FILE6_20.ITEM AS ITEM , FILE6_20.QUANT AS QUANT , FILE1_10.DESCA AS ITEMDESC " & _
                      "FROM FILE1_10 INNER JOIN FILE6_20 ON FILE1_10.ITEM = FILE6_20.ITEM " & _
                      " WHERE File6_20.Code = " & MyParn(ClientMove.xClient.Text) & _
                      " AND File6_20.DOC_NO = " & MyParn(xInv.Caption) & _
                      " ORDER BY FILE6_20.ITEM"
        Case 5
            Me.Caption = " „—œÊœ „»Ì‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹⁄«  "
            cString = "SELECT FILE6_10.STORE AS STORE , FILE6_10.TOTAL AS TOTAL , FILE6_10.PRICE AS PRICE , FILE6_10.ITEM AS ITEM , FILE6_10.QUANT AS QUANT , FILE1_10.DESCA AS ITEMDESC " & _
                      "FROM FILE1_10 INNER JOIN FILE6_10 ON FILE1_10.ITEM = FILE6_10.ITEM " & _
                      " WHERE File6_10.Code = " & MyParn(ClientMove.xClient.Text) & _
                      " AND File6_10.DOC_NO = " & MyParn(xInv.Caption) & _
                      " ORDER BY FILE6_10.ITEM"
    End Select
Else
    Select Case xType.Caption
        Case 4
            Me.Caption = " „‘ —Ì«  "
            cString = "SELECT FILE7_20.STORE AS STORE , FILE7_20.TOTAL AS TOTAL , FILE7_20.PRICE AS PRICE , FILE7_20.ITEM AS ITEM , FILE7_20.QUANT AS QUANT , FILE1_10.DESCA AS ITEMDESC " & _
                      "FROM FILE1_10 INNER JOIN FILE7_20 ON FILE1_10.ITEM = FILE7_20.ITEM " & _
                      " WHERE File7_20.Code = " & MyParn(ClientMove.xClient.Text) & _
                      " AND File7_20.DOC_NO = " & MyParn(xInv.Caption) & _
                      " ORDER BY FILE7_20.ITEM"
        Case 5
            Me.Caption = " „—œÊœ "
            cString = "SELECT FILE6_11.STORE AS STORE , FILE6_11.TOTAL AS TOTAL , FILE6_11.PRICE AS PRICE , FILE6_11.ITEM AS ITEM , FILE6_11.QUANT AS QUANT , FILE1_10.DESCA AS ITEMDESC " & _
                      "FROM FILE1_10 INNER JOIN FILE6_11 ON FILE1_10.ITEM = FILE6_11.ITEM " & _
                      " WHERE File6_11.Code = " & MyParn(ClientMove.xClient.Text) & _
                      " AND File6_11.DOC_NO = " & MyParn(xInv.Caption) & _
                      " ORDER BY FILE6_11.ITEM"
    End Select
End If
Set InvTable = mydb.CreateSnapshot(cString)
InvGrid.FormatString = "«·„Œ“‰|" & "ﬂÊœ «·’‰›|" & "«·’‰‹‹‹‹›|" & "«·ﬂ„Ì…|" & "«·”⁄—|" & "«·≈Ã„«·Ï"
InvGrid.Cols = 6
InvGrid.ColWidth(0) = 500
InvGrid.ColWidth(1) = 1500
InvGrid.ColWidth(2) = 3000
InvGrid.ColWidth(3) = 1200
InvGrid.ColWidth(4) = 1200
InvGrid.ColWidth(5) = 1200
If InvTable.RecordCount > 0 Then
    Fillgrd
Else
    InvGrid.Rows = 1
End If
End Sub

