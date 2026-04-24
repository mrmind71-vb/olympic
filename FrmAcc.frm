VERSION 5.00
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form FrmAcc 
   Caption         =   "Õ”«» «· ‘€Ì· Ê «·„ «Ã—…"
   ClientHeight    =   2235
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8310
   LinkTopic       =   "Form2"
   RightToLeft     =   -1  'True
   ScaleHeight     =   2235
   ScaleWidth      =   8310
   StartUpPosition =   3  'Windows Default
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      RightToLeft     =   -1  'True
      Top             =   240
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox xDate1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      Height          =   315
      Left            =   5625
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   682
      Width           =   1290
   End
   Begin VB.TextBox xDate2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      Height          =   315
      Left            =   5625
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   1125
      Width           =   1290
   End
   Begin VB.CommandButton Cmd_Ok 
      Caption         =   "⁄—÷ Õ/ «· ‘€Ì· Ê «·„ «Ã—…"
      Height          =   690
      Left            =   720
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   960
      Width           =   2340
   End
   Begin Crystal.CrystalReport Report1 
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
   Begin MSDBCtls.DBCombo xStore 
      Bindings        =   "FrmAcc.frx":0000
      DataSource      =   "Data1"
      Height          =   315
      Left            =   4560
      TabIndex        =   0
      Top             =   240
      Width           =   2355
      _ExtentX        =   4154
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
      Caption         =   "„Œ“‰"
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
      Left            =   7200
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   270
      Width           =   450
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "„‹‹‹‰  «—ÌŒ "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   7080
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   720
      Width           =   915
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "≈·‹‹Ï  «—Ì‹‹Œ"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   7140
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   1125
      Width           =   1260
   End
End
Attribute VB_Name = "FrmAcc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ChargTable As Recordset
Dim IncomTable As Recordset

Dim PurchTable As Recordset
Dim RetPurchTable As Recordset

Dim SalTable As Recordset
Dim RetSalTable  As Recordset

Dim Trans1Table As Recordset
Dim Trand2Table  As Recordset

Dim DamTable As Recordset
Dim InTable  As Recordset
Dim OutTable  As Recordset

Dim TempTable As Recordset
Dim Temp2Table As Recordset

Private Sub Cmd_OK_Click()
    Dim n1, n2, n3, n4, n5, n6, n7, n8 As Double
    Dim n10, N11 As Double
    Dim nProft1, nProft2 As Double
    Dim cHead As String
    Dim cHead2 As String
        
    If Not IsDate(xDate1.Text) Or Not IsDate(xDate2.Text) Then Exit Sub
    
    LTRUE = False
    
    cHead = " „‰  «—ÌŒ  " & Format(xDate1.Text, "DD-MM-YYYY") & " ≈·Ï  «—ÌŒ  " & Format(xDate2.Text, "DD-MM-YYYY")
    
    n1 = n2 = n3 = n4 = n5 = n6 = n7 = n8 = 0
    n10 = N11 = 0
    nProft1 = nProft2 = 0
    
    n1 = EvalBalStore(xDate1.Text, 1)
    n5 = EvalBalStore(xDate2.Text, 2)
    
    tempdb.Execute "DELETE * FROM TEMP"
    Set TempTable = tempdb.OpenRecordset("Temp")
    
    cString = "SELECT FILE8_70.DESCA, Sum(FILE8_50.VALUE) AS t_sum " & _
               " FROM (FILE8_50 LEFT JOIN FILE0_50 ON FILE8_50.BOX = FILE0_50.CODE) LEFT JOIN FILE8_70 ON FILE8_50.CHARGE = FILE8_70.CODE " & _
               " WHERE FILE8_50.[Date] Between DateValue(" & MyParn(xDate1.Text) & ")" & _
               " and  DateValue(" & MyParn(xDate2.Text) & ")" & _
               " and store = " & MyParn(xStore.BoundText) & _
               " GROUP BY FILE8_70.DESCA, FILE8_70.CODE "
    Set ChargTable = mydb.OpenRecordset(cString, dbOpenSnapshot)
    
    cString = "SELECT FILE8_71.DESCA, Sum(FILE8_60.VALUE) AS t_sum " & _
               " FROM (FILE8_60 LEFT JOIN FILE0_50 ON FILE8_60.BOX = FILE0_50.CODE) LEFT JOIN FILE8_71 ON FILE8_60.CHARGE = FILE8_71.CODE " & _
               " WHERE FILE8_60.[Date] Between DateValue(" & MyParn(xDate1.Text) & ")" & _
               " and  DateValue(" & MyParn(xDate2.Text) & ")" & _
               " and store = " & MyParn(xStore.BoundText) & _
               " GROUP BY FILE8_71.DESCA, FILE8_71.CODE "
    Set IncomTable = mydb.OpenRecordset(cString, dbOpenSnapshot)
    


    cStr1 = " SELECT Sum(FILE7_20.Total) AS T_Total" & _
            " FROM FILE7_20  " & _
            " WHERE STORE = " & MyParn(xStore.BoundText) & _
            " AND Date >= DateValue(" & MyParn(xDate1.Text) & ")" & _
            " and Date <= DateValue(" & MyParn(xDate2.Text) & ")"
    Set PurchTable = mydb.OpenRecordset(cStr1)

    cStr1 = " SELECT Sum(FILE6_11.Total) AS T_Total" & _
            " FROM FILE6_11  " & _
            " WHERE STORE = " & MyParn(xStore.BoundText) & _
            " AND Date >= DateValue(" & MyParn(xDate1.Text) & ")" & _
            " and Date <= DateValue(" & MyParn(xDate2.Text) & ")"
    Set RetPurchTable = mydb.OpenRecordset(cStr1)
    
    cStr1 = " SELECT Sum(FILE6_20.Total) AS T_Total" & _
            " FROM FILE6_20  " & _
            " WHERE STORE = " & MyParn(xStore.BoundText) & _
            " and Date >= DateValue(" & MyParn(xDate1.Text) & ")" & _
            " and Date <= DateValue(" & MyParn(xDate2.Text) & ")"
    Set SalTable = mydb.OpenRecordset(cStr1)

    cStr1 = " SELECT Sum(FILE6_10.Total) AS T_Total" & _
            " FROM FILE6_10  " & _
            " WHERE STORE = " & MyParn(xStore.BoundText) & _
            " AND Date >= DateValue(" & MyParn(xDate1.Text) & ")" & _
            " and Date <= DateValue(" & MyParn(xDate2.Text) & ")"
    Set RetSalTable = mydb.OpenRecordset(cStr1)
    
    
    cStr1 = " SELECT Sum(FILE1_60.TOTAL) AS t_TOTAL " & _
            " FROM FILE1_60  " & _
            " WHERE file1_60.store2 = " & MyParn(xStore.BoundText) & _
            " and Date >= DateValue(" & MyParn(xDate1.Text) & ")" & _
            " and Date <= DateValue(" & MyParn(xDate2.Text) & ")"
    Set Trans1Table = mydb.OpenRecordset(cStr1)
    
    cStr1 = " SELECT Sum(FILE1_60.TOTAL) AS t_TOTAL " & _
            " FROM FILE1_60 " & _
            " WHERE file1_60.store1 = " & MyParn(xStore.BoundText) & _
            " and Date >= DateValue(" & MyParn(xDate1.Text) & ")" & _
            " and Date <= DateValue(" & MyParn(xDate2.Text) & ")"
    Set Trans2Table = mydb.OpenRecordset(cStr1)
    
    cStr1 = " SELECT Sum(FILE1_80.Total) AS T_Total" & _
            " FROM FILE1_80  " & _
            " WHERE STORE = " & MyParn(xStore.BoundText) & _
            " and Date >= DateValue(" & MyParn(xDate1.Text) & ")" & _
            " and Date <= DateValue(" & MyParn(xDate2.Text) & ")"
    Set InTable = mydb.OpenRecordset(cStr1)
    
    cStr1 = " SELECT Sum(FILE1_81.Total) AS T_Total" & _
            " FROM FILE1_81  " & _
            " WHERE STORE = " & MyParn(xStore.BoundText) & _
            " and Date >= DateValue(" & MyParn(xDate1.Text) & ")" & _
            " and Date <= DateValue(" & MyParn(xDate2.Text) & ")"
    Set OutTable = mydb.OpenRecordset(cStr1)
    
    cStr1 = " SELECT Sum(FILE1_82.Total) AS T_Total" & _
            " FROM FILE1_82  " & _
            " WHERE STORE = " & MyParn(xStore.BoundText) & _
            " and Date >= DateValue(" & MyParn(xDate1.Text) & ")" & _
            " and Date <= DateValue(" & MyParn(xDate2.Text) & ")"
    Set DamTable = mydb.OpenRecordset(cStr1)
    
      
    PurchTable.MoveFirst
    n2 = TurnValue(PurchTable.t_total, Null, 0)
    RetPurchTable.MoveFirst
    n2 = n2 - TurnValue(RetPurchTable.t_total, Null, 0)
    
    
    SalTable.MoveFirst
    n6 = TurnValue(SalTable.t_total, Null, 0)
    RetSalTable.MoveFirst
    n6 = n6 - TurnValue(RetSalTable.t_total, Null, 0)
    
    Trans1Table.MoveFirst
    n3 = TurnValue(Trans1Table.t_total, Null, 0)
    
    Trans2Table.MoveFirst
    n7 = TurnValue(Trans2Table.t_total, Null, 0)
    
    InTable.MoveFirst
    n4 = TurnValue(InTable.t_total, Null, 0)
    
    OutTable.MoveFirst
    n8 = TurnValue(OutTable.t_total, Null, 0)
    
    DamTable.MoveFirst
    n8 = n8 + TurnValue(DamTable.t_total, Null, 0)
    
    With TempTable
    .AddNew
    .str18 = " Õ/ «·„ «Ã—…  " & xStore.Text
    .str19 = cHead
    
    .str1 = "»÷«⁄… √Ê· «·„œ… "
    .str2 = "»÷«⁄… √Œ— «·„œ… "
    .VAL1 = n1
    .VAL2 = n5
    .Update
    .AddNew
    .str18 = " Õ/ «·„ «Ã—…  " & xStore.Text
    .str19 = cHead
    .Update
    
    .AddNew
    .str18 = " Õ/ «·„ «Ã—…  " & xStore.Text
    .str19 = cHead
    .str1 = "„‘ —Ì«   "
    .str2 = "„»Ì⁄«  "
    .VAL1 = n2
    .VAL2 = n6
    .Update
    .AddNew
    .str18 = " Õ/ «·„ «Ã—…  " & xStore.Text
    .str19 = cHead
    .Update
    
    .AddNew
    .str18 = " Õ/ «·„ «Ã—…  " & xStore.Text
    .str19 = cHead
    .str1 = " ÕÊÌ·«  ≈·Ï «·„Õ·  "
    .str2 = " ÕÊÌ·«  „‹‰ «·„Õ·  "
    .VAL1 = n3
    .VAL2 = n7
    .Update
    .AddNew
    .str18 = " Õ/ «·„ «Ã—…  " & xStore.Text
    .str19 = cHead
    .Update
    
    .AddNew
    .str18 = " Õ/ «·„ «Ã—…  " & xStore.Text
    .str19 = cHead
    .str1 = "Ê«—œ ≈·Ï «·„Õ·  "
    .str2 = " ’«œ— & Â«·ﬂ „‰ «·„Õ· "
    .VAL1 = n4
    .VAL2 = n8
    .Update
    .AddNew
    .str18 = " Õ/ «·„ «Ã—…  " & xStore.Text
    .str19 = cHead
    .Update
    
    If (n5 + n6 + n7 + n8) - (n1 + n2 + n3 + n4) > 0 Then
        nProft1 = (n5 + n6 + n7 + n8) - (n1 + n2 + n3 + n4)
        .AddNew
        .str1 = "„Ã„· «·—»Õ   "
        .VAL1 = nProft1
        nProft2 = 0
        .Update
        .AddNew
        .Update
    Else
        nProft2 = (n1 + n2 + n3 + n4) - (n5 + n6 + n7 + n8)
        .AddNew
        .str2 = "„Ã„· «·Œ”«—…  "
        .VAL2 = nProft2
        nProft1 = 0
        .Update
        .AddNew
        .Update
    End If
    End With

    Report1.ReportFileName = PublicPath & "\Reports\R_Acc1.rpt"
    Report1.DataFiles(0) = App.Path & "\Temp.mdb"
    Report1.Action = 1

'*******************
'*******************
'*******************
    
    tempdb.Execute "DELETE * FROM TEM2"
    Set Temp2Table = tempdb.OpenRecordset("Tem2")
    With Temp2Table
    
    n10 = 0
    
    If ChargTable.RecordCount > 0 Then
        ChargTable.MoveFirst
        Do While Not ChargTable.EOF
            n10 = n10 + ChargTable.T_SUM
            I = I + 1
            .AddNew
            .str18 = " Õ/ «·√—»«Õ Ê «·Œ”«∆—   " & xStore.Text
            .str19 = cHead
            .str1 = ChargTable.DESCA
            .VAL1 = ChargTable.T_SUM
            .Update
            
            ChargTable.MoveNext
            If ChargTable.EOF Then Exit Do
            
            .AddNew
            .str18 = " Õ/ «·√—»«Õ Ê «·Œ”«∆—   " & xStore.Text
            .str19 = cHead
            .Update
        Loop
    End If
   
    If IncomTable.RecordCount > 0 Then
        IncomTable.MoveFirst
        Do While Not IncomTable.EOF
            N11 = N11 + IncomTable.T_SUM
            I = I + 1
            .AddNew
            .str18 = " Õ/ «·√—»«Õ Ê «·Œ”«∆—   " & xStore.Text
            .str19 = cHead
            .str2 = IncomTable.DESCA
            .VAL2 = IncomTable.T_SUM
            .Update
            
            IncomTable.MoveNext
            If IncomTable.EOF Then Exit Do
            
            .AddNew
            .str18 = " Õ/ «·√—»«Õ Ê «·Œ”«∆—   " & xStore.Text
            .str19 = cHead
            .Update
        Loop
    End If
       
    If nProft1 > 0 Then
        .AddNew
        .str18 = " Õ/ «·√—»«Õ Ê «·Œ”«∆—   " & xStore.Text
        .str19 = cHead
        .str2 = "„Ã„· «·—»Õ "
        .VAL2 = nProft1
        .Update
    End If
    
    If nProft2 > 0 Then
        .AddNew
        .str18 = " Õ/ «·√—»«Õ Ê «·Œ”«∆—   " & xStore.Text
        .str19 = cHead
        .str8 = cHead
        .str1 = "„Ã„· «·Œ”«—… "
        .VAL1 = nProft2
        .Update
        n10 = n10 + nProft2
    End If
    
    .AddNew
    .str18 = " Õ/ «·√—»«Õ Ê «·Œ”«∆—   " & xStore.Text
    .str19 = cHead
    .Update
    
'    If nProft1 > nProft2 Then
        If nProft1 < 0 Then
            If N11 > n10 Then
                .AddNew
                .str1 = "’«›Ï «·—»Õ "
                .VAL1 = N11 - n10
                .Update
            Else
                .AddNew
                .str2 = "’«›Ï «·Œ”«—… "
                .VAL2 = n10 - N11
                .Update
            End If
        
        Else
            If N11 + nProft1 > n10 Then
                .AddNew
                .str1 = "’«›Ï «·—»Õ "
                .VAL1 = N11 + nProft1 - n10
                .Update
            Else
                .AddNew
                .str2 = "’«›Ï «·Œ”«—… "
                .VAL2 = n10 - N11 - nProft1
                .Update
            End If
        End If
'    Else
'        .AddNew
'        .str2 = "’«›Ï «·Œ”«—… "
'        .VAL2 = (n10 - N11) + nProft2
'        .Update
'    End If
    
    .AddNew
    .Update
    
    End With
    
    Report1.ReportFileName = PublicPath & "\Reports\R_Acc2.rpt"
    Report1.DataFiles(0) = App.Path & "\Temp.mdb"
    Report1.Action = 1

End Sub
Private Sub Form_Load()
    Data1.DatabaseName = MdbPath
    Data1.RecordSource = "SELECT CODE , DESCA FROM FILE1_70 WHERE FLAG = 1 "
    xStore.ListField = "Desca"
    xStore.BoundColumn = "code"

End Sub
Private Sub VsSal_EnterCell()
    With VsSal
        If .Col = 3 Then
            .Editable = flexEDKbdMouse
        Else
            .Editable = flexEDNone
        End If
    End With
End Sub
Private Function EvalBalStore(dDate, nType)
Dim LastCost As Recordset
Dim SourceTable As Recordset
Dim nTot As Double
Dim nBal As Double
If nType = 1 Then
    cStr1 = " SELECT FILE7_20.ITEM, FILE7_20.Price, FILE7_20.DATE FROM FILE7_20  " & _
            " WHERE Date < DateValue(" & MyParn(dDate) & ")" & _
            " ORDER BY DATE "
Else
    cStr1 = " SELECT FILE7_20.ITEM, FILE7_20.Price, FILE7_20.DATE FROM FILE7_20  " & _
            " WHERE Date <= DateValue(" & MyParn(dDate) & ")" & _
            " ORDER BY DATE "
End If
Set LastCost = mydb.OpenRecordset(cStr1)

cString = "SELECT FILE1_10.ITEM, First(FILE1_10.PRICE) AS ItemPRICE, " & _
          " First(FILE1_10.COST) AS ItemCOST, First(FILE1_10.DESCA) AS ItemDESCA,  " & _
          "Sum(FILE1_11.[IN]) AS SumIN, Sum(FILE1_11.OUT) AS SumOUT, " & _
          " First(FILE1_50.M_GROUP) AS MGrCode, First(FILE1_50.DESCA) AS GrDESC,  " & _
          " First(FILE1_10.GROUP) AS GrCode, First(FILE1_70.DESCA) AS MGrDesc" & _
          " FROM ((FILE1_10 RIGHT JOIN FILE1_11 ON FILE1_10.ITEM = FILE1_11.ITEM) LEFT " & _
          " JOIN FILE1_50 ON FILE1_10.GROUP = FILE1_50.CODE) LEFT JOIN FILE1_70 ON " & _
          " FILE1_50.M_GROUP = FILE1_70.CODE"

sQuery = " Where File1_70.Flag = 2"
sQuery = myQuery(sQuery) & "File1_11.store = " & MyParn(xStore.BoundText)
If nType = 1 Then
    sQuery = myQuery(sQuery) & " file1_11.Date < DateValue(" & MyParn(dDate) & ")"
Else
    sQuery = myQuery(sQuery) & " file1_11.Date <= DateValue(" & MyParn(dDate) & ")"
End If
cString = cString & sQuery & " GROUP BY File1_10.ITEM "
Set SourceTable = mydb.OpenRecordset(cString, dbOpenSnapshot)
nTot = 0
With SourceTable
SourceTable.MoveFirst
If .RecordCount > 0 Then
    Do While Not .EOF
        If Format((TurnValue(SourceTable.SumIn, Null, 0) - TurnValue(SourceTable.SUMout, Null, 0)), "##0.00") <> "0.00" Then
            Me.Cmd_Ok.Caption = .Item
            nBal = TurnValue(.SumIn, Null, 0) - TurnValue(.SUMout, Null, 0)
            LastCost.FindLast " ITEM = " & MyParn(.Item)
            If LastCost.NoMatch Then
                nTot = nTot + (.ITEMcost * nBal)
            Else
                nTot = nTot + (LastCost.price * nBal)
            End If
        End If
       .MoveNext
    Loop
End If
End With
EvalBalStore = nTot
End Function


