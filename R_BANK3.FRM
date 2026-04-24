VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form R_Bank3 
   Caption         =   " ﬁ«—Ì— «·»‰Êﬂ"
   ClientHeight    =   1860
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4740
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   RightToLeft     =   -1  'True
   ScaleHeight     =   1860
   ScaleWidth      =   4740
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox xDate2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   390
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   240
      Width           =   1155
   End
   Begin VB.TextBox xdate1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2430
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   240
      Width           =   1230
   End
   Begin Crystal.CrystalReport Report1 
      Left            =   4035
      Top             =   1215
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
   Begin VB.Frame Frame1 
      Height          =   645
      Left            =   165
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   1035
      Width           =   2385
      Begin VB.CommandButton CmdExit 
         Caption         =   "Œ—ÊÃ"
         Height          =   375
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   180
         Width           =   1095
      End
      Begin VB.CommandButton CmdApply 
         Caption         =   "⁄—÷"
         Height          =   375
         Left            =   1275
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   180
         Width           =   1005
      End
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "«·Ï :"
      Height          =   225
      Left            =   1635
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   285
      Width           =   390
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
      Left            =   3750
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   195
      Width           =   765
   End
End
Attribute VB_Name = "R_Bank3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TempTable As Recordset
Dim BankTable As Recordset
Dim nOption As Integer
Dim cHead As String
Dim cF_BAL As String, FBalTable As Recordset
Dim cC_BAL As String, CBalTable As Recordset
Dim cT_In As String, InTable As Recordset
Dim cT_Out As String, OutTable As Recordset
Dim cT_Charg As String, ChargTable As Recordset
Dim cChq_In As String, ChqInTable As Recordset
Dim cChq_Out As String, ChqOutTable As Recordset
Dim cChq1 As String, cChq2 As String, Chq1Table As Recordset, Chq2Table As Recordset
Function MYVALID()
If Not IsDate(xdate1.Text) Then Exit Function
If Not IsDate(xDate2.Text) Then Exit Function
MYVALID = True
End Function
Private Sub CmdExit_Click()
Unload Me
End Sub
Private Sub CmdUndo_Click()
xdate1.Text = ""
xDate2.Text = ""
End Sub
Private Sub CmdApply_Click()
If Not MYVALID Then Exit Sub

tempdb.Execute "Delete * From Temp"
Set TempTable = tempdb.OpenRecordset("TEMP")
Set BankTable = mydb.OpenRecordset("SELECT FILE5_10.* FROM FILE5_10 ORDER BY FILE5_10.ID ")

cT_In = "SELECT file5_11.BANK, Sum(file5_11.[IN]) AS SumIN FROM file5_11 " & _
       " WHERE file5_11.TYPE = '2' " & _
       " AND Date >= DateValue(" & MyParn(xdate1.Text) & ")" & _
       " and Date <= DateValue(" & MyParn(xDate2.Text) & ")" & _
       " GROUP BY file5_11.BANK "

cT_Out = "SELECT file5_11.BANK, Sum(file5_11.[OUT]) AS SumOut FROM file5_11 " & _
       " WHERE file5_11.TYPE = '3' " & _
       " AND Date >= DateValue(" & MyParn(xdate1.Text) & ")" & _
       " and Date <= DateValue(" & MyParn(xDate2.Text) & ")" & _
       " GROUP BY file5_11.BANK "

cT_Charg = "SELECT file5_11.BANK, Sum(file5_11.[OUT]) AS SumCharg FROM file5_11 " & _
       " WHERE file5_11.TYPE = '4' " & _
       " AND Date >= DateValue(" & MyParn(xdate1.Text) & ")" & _
       " and Date <= DateValue(" & MyParn(xDate2.Text) & ")" & _
       " GROUP BY file5_11.BANK "

cChq_In = "SELECT file5_11.BANK, Sum(file5_11.[IN]) AS SumChqIn FROM file5_11 " & _
       " WHERE file5_11.TYPE = '5' " & _
       " AND Date >= DateValue(" & MyParn(xdate1.Text) & ")" & _
       " and Date <= DateValue(" & MyParn(xDate2.Text) & ")" & _
       " GROUP BY file5_11.BANK "

cChq_Out = "SELECT file5_11.BANK, Sum(file5_11.[OUT]) AS SumChqOut FROM file5_11 " & _
       " WHERE file5_11.TYPE = '6' " & _
       " AND Date >= DateValue(" & MyParn(xdate1.Text) & ")" & _
       " and Date <= DateValue(" & MyParn(xDate2.Text) & ")" & _
       " GROUP BY file5_11.BANK "

cF_BAL = "SELECT file5_11.BANK, Sum(file5_11.[IN]) AS SumIN, Sum(file5_11.OUT) AS SumOUT FROM file5_11" & _
       " WHERE Date < DateValue(" & MyParn(xdate1.Text) & ")" & _
       " GROUP BY file5_11.BANK "

cC_BAL = "SELECT file5_11.BANK, Sum(file5_11.[IN]) AS SumIN, Sum(file5_11.OUT) AS SumOUT FROM file5_11" & _
       " WHERE Date <= DateValue(" & MyParn(xDate2.Text) & ")" & _
       " GROUP BY file5_11.BANK "

cChq1 = "SELECT Sum(FILE5_20.VALUE) AS SumChq, FILE5_20.ID_BANK FROM FILE5_20 " & _
       " WHERE FILE5_20.CLOSE ='0' " & _
       " GROUP BY FILE5_20.ID_BANK "

cChq2 = "SELECT Sum(FILE5_21.VALUE) AS SumChq, FILE5_21.ID_BANK FROM FILE5_21 " & _
       " WHERE FILE5_21.CLOSE ='0' " & _
       " GROUP BY FILE5_21.ID_BANK "

Set InTable = mydb.OpenRecordset(cT_In)
Set OutTable = mydb.OpenRecordset(cT_Out)
Set ChargTable = mydb.OpenRecordset(cT_Charg)
Set ChqInTable = mydb.OpenRecordset(cChq_In)
Set ChqOutTable = mydb.OpenRecordset(cChq_Out)
Set FBalTable = mydb.OpenRecordset(cF_BAL)
Set CBalTable = mydb.OpenRecordset(cC_BAL)
Set Chq1Table = mydb.OpenRecordset(cChq1)
Set Chq2Table = mydb.OpenRecordset(cChq2)

With BankTable
    If .RecordCount > 0 Then
        .MoveFirst
        Do While Not .EOF
            TempTable.AddNew
            TempTable.str1 = .ID
            TempTable.str2 = .DESCA
            TempTable.str4 = " ≈Ã„«·Ï  ⁄«„·«  «·»‰Êﬂ Œ·«· › —… "
            TempTable.STR5 = " „‰  «—ÌŒ " & xdate1.Text & " ≈·Ï  «—ÌŒ " & xDate2.Text
            
            FBalTable.FindFirst " BANK = " & MyParn(.ID)
            If Not FBalTable.NoMatch Then TempTable.val1 = TurnValue(FBalTable.SumIn, Null, 0) - TurnValue(FBalTable.SUMout, Null, 0)
            
            InTable.FindFirst " BANK = " & MyParn(.ID)
            If Not InTable.NoMatch Then TempTable.VAL2 = TurnValue(InTable.SumIn, Null, 0)
            
            OutTable.FindFirst " BANK = " & MyParn(.ID)
            If Not OutTable.NoMatch Then TempTable.VAL3 = OutTable.SUMout
            
            ChargTable.FindFirst " BANK = " & MyParn(.ID)
            If Not ChargTable.NoMatch Then TempTable.VAL4 = ChargTable.SUMCHARG
            
            ChqInTable.FindFirst " BANK = " & MyParn(.ID)
            If Not ChqInTable.NoMatch Then TempTable.VAL5 = ChqInTable.SumChqIn
            
            ChqOutTable.FindFirst " BANK = " & MyParn(.ID)
            If Not ChqOutTable.NoMatch Then TempTable.VAL6 = ChqOutTable.SumChqOUT
            
            CBalTable.FindFirst " BANK = " & MyParn(.ID)
            If Not CBalTable.NoMatch Then TempTable.VAL7 = TurnValue(CBalTable.SumIn, Null, 0) - TurnValue(CBalTable.SUMout, Null, 0)
            
            Chq1Table.FindFirst " ID_BANK = " & MyParn(.ID)
            If Not Chq1Table.NoMatch Then TempTable.VAL8 = TurnValue(CBalTable.SumIn, Null, 0) - TurnValue(Chq1Table.SUMCHQ, Null, 0)
            
            Chq2Table.FindFirst " ID_BANK = " & MyParn(.ID)
            If Not Chq2Table.NoMatch Then TempTable.VAL9 = TurnValue(CBalTable.SumIn, Null, 0) - TurnValue(Chq2Table.SUMCHQ, Null, 0)
            
            TempTable.str19 = firsttitle
            TempTable.str20 = Secondtitle
            TempTable.Update
            .MoveNext
        Loop
    End If
End With
Report1.ReportFileName = PublicPath & "\Reports\BANKREP3.RPT"
Report1.DataFiles(0) = App.Path & "\temp.mdb"
Report1.Action = 1
End Sub
Function TurnValue(pSource, pOld, pNew)
   TurnValue = IIf(pSource = pOld Or (IsNull(pSource) And IsNull(pOld)), pNew, pSource)
End Function
