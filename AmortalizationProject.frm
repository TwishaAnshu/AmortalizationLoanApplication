VERSION 5.00
Begin VB.Form AmortalizationProject 
   Caption         =   "Form1"
   ClientHeight    =   9732
   ClientLeft      =   120
   ClientTop       =   456
   ClientWidth     =   17556
   LinkTopic       =   "Form1"
   ScaleHeight     =   9732
   ScaleWidth      =   17556
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtYearlyExtra 
      Height          =   1032
      Left            =   10140
      TabIndex        =   26
      Top             =   4500
      Width           =   1512
   End
   Begin VB.TextBox txtMonthly 
      Height          =   972
      Left            =   6060
      TabIndex        =   25
      Top             =   4620
      Width           =   1512
   End
   Begin VB.TextBox txtOneTimeExtra 
      Height          =   972
      Left            =   8340
      TabIndex        =   24
      Top             =   4560
      Width           =   1272
   End
   Begin VB.Frame Frame 
      Caption         =   "Amortalization Table"
      Height          =   1815
      Index           =   1
      Left            =   840
      TabIndex        =   15
      Top             =   7080
      Width           =   12495
      Begin VB.HScrollBar hsbPayment 
         Height          =   255
         Left            =   480
         TabIndex        =   17
         Top             =   1320
         Width           =   11655
      End
      Begin VB.TextBox txtAmortTable 
         Height          =   735
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   16
         Top             =   540
         Width           =   11775
      End
      Begin VB.Label lblMONTHLYPTINCIPle 
         Caption         =   "monthly principle"
         Height          =   252
         Left            =   8220
         TabIndex        =   23
         Top             =   240
         Width           =   1452
      End
      Begin VB.Label lblMONTHLYINTERST 
         Caption         =   "monthly interest"
         Height          =   312
         Left            =   6660
         TabIndex        =   22
         Top             =   300
         Width           =   1452
      End
      Begin VB.Label Label 
         Caption         =   "total interest"
         Height          =   192
         Index           =   6
         Left            =   5340
         TabIndex        =   21
         Top             =   300
         Width           =   2052
      End
      Begin VB.Label Label 
         Caption         =   "current balance"
         Height          =   252
         Index           =   5
         Left            =   3240
         TabIndex        =   20
         Top             =   300
         Width           =   1992
      End
      Begin VB.Label Year 
         Caption         =   "Year"
         Height          =   312
         Left            =   2280
         TabIndex        =   19
         Top             =   300
         Width           =   1092
      End
      Begin VB.Label Label 
         Caption         =   "payment#"
         Height          =   195
         Index           =   2
         Left            =   360
         TabIndex        =   18
         Top             =   300
         Width           =   1995
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "MONTHLY PAYMENT"
      Height          =   2655
      Left            =   720
      TabIndex        =   13
      Top             =   4140
      Width           =   3255
      Begin VB.Label lblMonthlyPayment 
         Height          =   1095
         Left            =   300
         TabIndex        =   14
         Top             =   840
         Width           =   2535
      End
   End
   Begin VB.Frame Frame 
      Caption         =   "Enter Values"
      Height          =   2715
      Index           =   0
      Left            =   180
      TabIndex        =   9
      Top             =   600
      Width           =   10155
      Begin VB.TextBox txtYears 
         Height          =   1035
         Left            =   6600
         TabIndex        =   2
         Top             =   960
         Width           =   2115
      End
      Begin VB.TextBox txtYearlyRate 
         Height          =   912
         Left            =   3240
         TabIndex        =   1
         Top             =   960
         Width           =   1992
      End
      Begin VB.TextBox txtLoanAmount 
         Height          =   855
         Left            =   420
         TabIndex        =   0
         Top             =   1020
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "YEARS"
         Height          =   855
         Left            =   6780
         TabIndex        =   12
         Top             =   420
         Width           =   2235
      End
      Begin VB.Label Label 
         Caption         =   "YEARLY INTEREST RATE"
         Height          =   555
         Index           =   1
         Left            =   3060
         TabIndex        =   11
         Top             =   360
         Width           =   2835
      End
      Begin VB.Label Label 
         Caption         =   "LOAN AMOUNT"
         Height          =   435
         Index           =   0
         Left            =   300
         TabIndex        =   10
         Top             =   420
         Width           =   1995
      End
   End
   Begin VB.CommandButton cmdQUIT 
      Caption         =   "QUIT"
      Height          =   1092
      Left            =   12000
      TabIndex        =   5
      Top             =   5280
      Width           =   1692
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "CLEAR"
      Height          =   972
      Left            =   11940
      TabIndex        =   4
      Top             =   3960
      Width           =   1752
   End
   Begin VB.CommandButton cmdCalculate 
      Caption         =   "CALCULATE"
      Height          =   1035
      Left            =   11880
      TabIndex        =   3
      Top             =   2640
      Width           =   1815
   End
   Begin VB.Label Label 
      Caption         =   "YEARLY EXTRA"
      Height          =   732
      Index           =   4
      Left            =   9960
      TabIndex        =   8
      Top             =   3900
      Width           =   1332
   End
   Begin VB.Label MONTHLY 
      Caption         =   "MONTHLY EXTRA"
      Height          =   492
      Left            =   8100
      TabIndex        =   7
      Top             =   3780
      Width           =   1692
   End
   Begin VB.Label Label 
      Caption         =   "ONE TIME EXTRA"
      Height          =   252
      Index           =   3
      Left            =   5940
      TabIndex        =   6
      Top             =   3960
      Width           =   1452
   End
End
Attribute VB_Name = "AmortalizationProject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim LoanAmount, MonthlyPayment As Currency
Dim YearlyRate, MonthlyRate As Single
Dim Years, Payments As Integer
Dim AmortTable(360) As String
Dim PaymentNumber As Integer
Dim MonthlyInt, TotalInt, CurrentAmt As Currency
Dim YearNumber As Integer
Dim DispLine As String
Dim OneTimeExtra As Integer
Dim MonthyExtra As Integer
Dim YearlyExtra, monthlyinterest, monthlyprinciple As Integer






Private Sub cmdCalculate_Click()
'reading values from text
LoanAmount = Val(txtLoanAmount) - OneTimeExtra - YearlyExtra

YearlyRate = Val(txtYearlyRate)
Years = Val(txtYears)
OneTimeExtra = Val(txtOneTimeExtra)
MonthyExtra = Val(txtOneTimeExtra)
YearlyExtra = Val(txtYearlyExtra)

'intermediate calculations
MonthlyRate = YearlyRate / 1200
Payments = Years * 12

'monthlypayment
MonthlyPayment = (LoanAmount * MonthlyRate / (1 - (1 + MonthlyRate) ^ (-Payments))) + MonthyExtra


'displayResults
lblMonthlyPayment = Format$(MonthlyPayment, "Currency")

'vbTAB instead of constTbCh = Chr$(9) . payment 12 should still be in the first year
'vbtab
' initiliazes totalInterest and Current balance

TotalInt = 0
CurrentAmt = LoanAmount
'-set up loop
For PaymentNumber = 1 To Payments
'Make calculations
    monthlyinterest = CurrentAmt * MonthlyRate
    monthlyprinciple = (MonthlyPayment - MonthlyInt)
    
    MonthlyInt = CurrentAmt * MonthlyRate
    TotalInt = TotalInt + MonthlyInt
    CurrentAmt = CurrentAmt + MonthlyInt - MonthlyPayment
    YearNumber = Int(PaymentNumber \ 12) + 1
If PaymentNumber Mod 12 = 0 Then
    YearNumber = Int(YearNumber - 1)
    DispLine = DispLine & vbTab & Format$(YearNumber, "#0")
End If
'build display line
    DispLine = vbTab & Format$(PaymentNumber, "####")
    DispLine = DispLine & vbTab & vbTab & Format$(YearNumber, "#0")
    DispLine = DispLine & vbTab & vbTab & Format$(CurrentAmt, "Currency")
        If CurrentAmt = 0 Then
             hsbPayment.Enabled = False
        End If
    DispLine = DispLine & vbTab & vbTab & Format$(TotalInt, "Currency")
    DispLine = DispLine & vbTab & vbTab & Format$(monthlyinterest, "Currency")
    DispLine = DispLine & vbTab & vbTab & Format$(monthlyprinciple, "Currency")
'-transferdisplayline to array element
    AmortTable(PaymentNumber) = DispLine

Next PaymentNumber '-end loop
'set up scroll bar
hsbPayment.Min = 1
hsbPayment.Max = Payments
hsbPayment.LargeChange = 12 ' one year = 12months
hsbPayment.Value = 1
'-put first line of table in textbox
txtAmortTable = AmortTable(1)



End Sub

Private Sub cmdClear_Click()
lblMonthlyPayment = ""
txtLoanAmount.SetFocus

End Sub

Private Sub CMDQUIT_Click()
End
End Sub

Private Sub hsbPayment_Change()

txtAmortTable = AmortTable(hsbPayment.Value)








End Sub

