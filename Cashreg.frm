VERSION 5.00
Begin VB.Form CashReg 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " "
   ClientHeight    =   3795
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5775
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   3795
   ScaleWidth      =   5775
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      Caption         =   "«·›Ì“«"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   180
      RightToLeft     =   -1  'True
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   3060
      Width           =   915
   End
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   135
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   2250
      Width           =   5505
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         Caption         =   "€Ì— „”œœ :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   3420
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   225
         Width           =   1275
      End
      Begin VB.Label xLate 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   420
         Left            =   1125
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   180
         Width           =   2175
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2220
      Left            =   135
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   45
      Width           =   5505
      Begin VB.TextBox xPay 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   1125
         MaxLength       =   15
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Top             =   1170
         Width           =   2175
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         Caption         =   "«·»«ﬁÏ :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   420
         Left            =   3375
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   1710
         Width           =   1275
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         Caption         =   "«·„œ›Ê⁄ :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   3375
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   1215
         Width           =   1275
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         Caption         =   "≈Ã„«·Ï ﬁÌ„… «·»Ê‰ :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   3375
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   720
         Width           =   1905
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         Caption         =   "⁄œœ «·ﬁÿ⁄ «·„»«⁄… :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   3375
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   270
         Width           =   1905
      End
      Begin VB.Label xtotalQuant 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   420
         Left            =   1125
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   180
         Width           =   2175
      End
      Begin VB.Label xTotal 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   420
         Left            =   1125
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   675
         Width           =   2175
      End
      Begin VB.Label xRest 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   420
         Left            =   1125
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   1665
         Width           =   2175
      End
   End
   Begin VB.Frame fmVisa 
      Height          =   735
      Left            =   1260
      RightToLeft     =   -1  'True
      TabIndex        =   14
      Top             =   2970
      Visible         =   0   'False
      Width           =   4425
      Begin VB.TextBox xVisa 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   135
         MaxLength       =   15
         RightToLeft     =   -1  'True
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   225
         Width           =   2175
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         Caption         =   "›Ì“« :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   2385
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   270
         Width           =   645
      End
   End
End
Attribute VB_Name = "CashReg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim lExit As Boolean
Dim bSave As Boolean
Dim oSearchCr As New Search3
Public myform As Form
Private Sub Check1_Click()
If Check1.Value = 0 Then
'  xVisa.Text = ""
   fmVisa.Visible = False
'   CalcTotals
Else
    fmVisa.Visible = True
End If
End Sub

Private Sub cmdCur_Click()
Crlookup
End Sub
Private Sub Form_Load()
Me.Height = Me.Height - fmCur.Height
xTotal.Caption = Myvalue(myform.xTotal.Text, "Fixed")
xPay.Text = Myvalue(myform.xPay.Caption, "Fixed")
xtotalQuant.Caption = Myvalue(myform.xtotalQuant.Caption, "Fixed")
CalcTotals
'myForm.bSave = False
End Sub
Private Sub Form_Unload(Cancel As Integer)
myform.bSave = bSave
Set CashReg = Nothing
End Sub

Private Sub xPay_Change()
    CalcTotals
End Sub
Private Sub xPay_GotFocus()
    xPay.SelStart = 0
    xPay.SelLength = Len(xPay.Text)
End Sub
Private Sub xRet_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then Exit Sub
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then Unload Me
End Sub
Private Sub CalcTotals()
xRest.Tag = Format(Val(xPay.Text) + Val(xVisa.Text) - Val(xTotal.Caption), "#0.00")
xRest.Caption = IIf(Val(xRest.Tag) >= 0 And Val(xTotal.Caption) > 0, Val(xRest.Tag), "")
xLate.Tag = Format(Val(xTotal.Caption) - Val(xPay.Text) - Val(xVisa.Text), "#0.00")
xLate.Caption = IIf(Val(xLate.Tag) >= 0, Val(xLate.Tag), "")
If Val(xRate.Caption) <> 0 Then
    xTotal_cr.Caption = Myvalue(Round(Val(xTotal.Caption) / Val(xRate.Caption), 2))
End If
End Sub
Private Sub xPay_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If (IsNumeric(xPay.Text) Or xPay.Text = "") And Val(xTotal.Caption) > 0 And Val(xPay.Text) < Val(xTotal.Caption) And Check1.Value = 1 Then
        xVisa.SetFocus
    Else
        myHandle
    End If
End If
End Sub

Private Sub xPay_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
End If
End Sub
Private Sub xRate_Change()
CalcTotals
End Sub

Private Sub xVisa_Change()
    CalcTotals
End Sub
Private Sub xVisa_GotFocus()
'    If Val(xTotal.Caption) - Val(xPay.Text) > 0 Then
'        xVisa.Text = Format((Val(xTotal.Caption) - Val(xPay.Text)), "#0.00")
'    End If
End Sub
Private Sub xVisa_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then myHandle
End Sub
Private Sub myHandle()
If (IsNumeric(xPay.Text) Or IsNumeric(xVisa.Text)) Then
    KeyCode = 0
    If Not IsNumeric(xPay.Text) Then xPay.Text = ""
    If Val(xPay.Text) < 0 Then xPay.Text = ""
    
    If Not IsNumeric(xVisa.Text) Then xVisa.Text = ""
    If Val(xVisa.Text) < 0 Then xVisa.Text = ""
    
    If xPay.Text = "" And xVisa.Text = "" Then Exit Sub
    
    If Val(xTotal.Caption) <= 0 And (Val(xPay.Text) <> 0 Or Val(xVisa.Text) <> 0) Then
        Inform "ﬁÌ„… «·„œ›Ê⁄ ›Ï Õ«·… »Ê‰ «·„— Ã⁄ ÌÃ» «‰  ”«ÊÌ ’›—"
        Exit Sub
    End If
        
    If Val(xTotal.Caption) >= 0 And Val(xRest.Tag) >= 0 Then
        If Val(xVisa.Text) > Val(xTotal.Caption) Then
            MsgBox "ﬁÌ„… «·›Ì“« «ﬂ»— „‰ ﬁÌ„… «·»Ê‰"
            Exit Sub
        ElseIf Val(xVisa.Text) > 0 And Val(xPay.Text) > 0 Then
            If Val(xPay.Text) >= Val(xTotal.Caption) Then
                MsgBox "ﬁÌ„… «·‰ﬁœÌ «ﬂ»— „‰ ﬁÌ„… «·›Ì“« ÊÌÊÃœ ”œ«œ ›Ì“«"
                Exit Sub
            End If
        End If
    
        myform.xPay = Myvalue(xPay.Text)
        myform.xRest.Caption = xRest.Caption
        
        myform.xCash.Caption = xTotal.Caption - Val(xVisa.Text)
        myform.xVisa.Caption = xVisa.Text
        myform.xLate.Caption = ""
        myform.xcur.Caption = xcode_cr.Caption
        bSave = True
        Unload Me
    ElseIf Val(xTotal.Caption) < 0 Then
        myform.xPay = ""
        myform.xVisa = ""
        myform.xRest.Caption = ""
        myform.xCash.Caption = IIf(myform.xisCash.Value = 1, Val(xTotal.Caption), 0)
        myform.xLate.Caption = IIf(myform.xisCash.Value = 0, Val(xTotal.Caption), 0)
        myform.xcur.Caption = xcode_cr.Caption
        bSave = True
        Unload Me
    ElseIf Val(xLate.Tag) > 0 Then
        If myform.xisCash.Value = 0 Then
            If MsgBox("”œ«œ " & xPay.Text & " „‰ ﬁÌ„… «·›« Ê—… Ê «·»«ﬁÏ  " & Val(xLate.Tag) & " √Ã·", vbYesNo + vbDefaultButton2) = vbYes Then
                myform.xPay.Caption = xPay.Text
                myform.xCash.Caption = xPay.Text
                myform.xVisa.Caption = xVisa.Text
                myform.xLate.Caption = Val(xLate.Tag)
                myform.xRest.Caption = ""
                myform.xcur.Caption = xcode_cr.Caption
                bSave = True
                Unload Me
            Else
                xPay.Text = ""
            End If
        End If
    End If
End If
End Sub
Private Sub Crlookup()
Dim Generalarray(5)
Dim listarray(0, 5)
Dim GrdArray(2, 1)

Set Generalarray(0) = Me

Generalarray(1) = "Select code,DescA,Rate From cur_codes"
Generalarray(2) = "Order by code"
Generalarray(3) = 5000
Generalarray(5) = False

listarray(0, 0) = "«·»Ì«‰"
listarray(0, 1) = "(%%DESCA%%)"

GrdArray(0, 0) = "«·ﬂÊœ"
GrdArray(0, 1) = 1000

GrdArray(1, 0) = "«·«”„"
GrdArray(1, 1) = 5000

GrdArray(2, 0) = "”⁄— «·’—›"
GrdArray(2, 1) = 1000

searchArray = Array(Generalarray, listarray, GrdArray)
oSearchCr.Caption = "≈” ⁄·«„ «·⁄„·«¡"
oSearchCr.Show 1
End Sub
Sub myproc()
xcode_cr.Caption = oSearchCr.grid1.TextMatrix(oSearchCr.grid1.Row, 0)
xDescA.Caption = oSearchCr.grid1.TextMatrix(oSearchCr.grid1.Row, 1)
xRate.Caption = oSearchCr.grid1.TextMatrix(oSearchCr.grid1.Row, 2)
If fmCur.Visible = False Then
    Me.Height = Me.Height + fmCur.Height
    fmCur.Visible = True
End If
Unload oSearchCr
End Sub

