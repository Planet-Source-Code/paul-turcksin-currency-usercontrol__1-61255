VERSION 5.00
Begin VB.Form frmDemo 
   Caption         =   "Currency usercontrol demo"
   ClientHeight    =   8565
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6120
   Icon            =   "Demo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8565
   ScaleWidth      =   6120
   StartUpPosition =   2  'CenterScreen
   Begin DemoCurrency.ucCurrency ucCurrency1 
      Height          =   615
      Left            =   1680
      TabIndex        =   41
      Top             =   1440
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1085
      CurrSymbol      =   "$"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      LeadingZero     =   1
      ThousandsSep    =   ","
      ThousandsGroup  =   3
      Text            =   "0"
   End
   Begin VB.CommandButton cmdShow 
      Caption         =   "what's  returned to program?"
      Height          =   255
      Left            =   360
      TabIndex        =   40
      Top             =   3720
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   615
      Left            =   3600
      MultiLine       =   -1  'True
      TabIndex        =   39
      Text            =   "Demo.frx":08CA
      Top             =   1440
      Width           =   2295
   End
   Begin VB.Frame frMisc 
      Caption         =   "Miscellaneous "
      Height          =   1875
      Left            =   2880
      TabIndex        =   25
      Top             =   2160
      Width           =   3135
      Begin VB.CommandButton cmdApply 
         Caption         =   "Apply"
         Height          =   255
         Left            =   2160
         TabIndex        =   37
         Top             =   1560
         Width           =   735
      End
      Begin VB.TextBox fGroup 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2400
         TabIndex        =   36
         Text            =   "Text1"
         Top             =   1220
         Width           =   495
      End
      Begin VB.TextBox fThouSep 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2400
         TabIndex        =   35
         Text            =   "Text1"
         Top             =   960
         Width           =   495
      End
      Begin VB.TextBox fDecSep 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2400
         TabIndex        =   34
         Text            =   "Text1"
         Top             =   700
         Width           =   495
      End
      Begin VB.TextBox fDigits 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2400
         TabIndex        =   33
         Text            =   "Text1"
         Top             =   440
         Width           =   495
      End
      Begin VB.CheckBox chLeading 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "Leading zero"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   32
         Top             =   1440
         Width           =   1335
      End
      Begin VB.TextBox fCurrSymbol 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2400
         TabIndex        =   27
         Text            =   "Text1"
         Top             =   180
         Width           =   495
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "Fractional digits"
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   480
         Width           =   2055
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "Thousands grouping"
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   1200
         Width           =   2055
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "Thousands separator"
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   960
         Width           =   2055
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Decimal separator"
         Height          =   375
         Left            =   120
         TabIndex        =   28
         Top             =   720
         Width           =   2055
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Monetary symbol"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.Frame frNegFmt 
      Caption         =   "Negative value formats"
      Height          =   4455
      Left            =   240
      TabIndex        =   8
      Top             =   4080
      Width           =   5775
      Begin VB.OptionButton optNegFmt 
         Caption         =   "Left parenthesis, number, space, monetary symbol, right parenthesis"
         Height          =   255
         Index           =   15
         Left            =   240
         TabIndex        =   24
         Top             =   3960
         Width           =   5300
      End
      Begin VB.OptionButton optNegFmt 
         Caption         =   "Left parenthesis, monetary symbol, space, number, right parenthesis"
         Height          =   255
         Index           =   14
         Left            =   240
         TabIndex        =   23
         Top             =   3720
         Width           =   5300
      End
      Begin VB.OptionButton optNegFmt 
         Caption         =   "Number, negative sign, space, monetary symbol"
         Height          =   255
         Index           =   13
         Left            =   240
         TabIndex        =   22
         Top             =   3480
         Width           =   5300
      End
      Begin VB.OptionButton optNegFmt 
         Caption         =   " Monetary symbol, space, negative sign, number"
         Height          =   255
         Index           =   12
         Left            =   240
         TabIndex        =   21
         Top             =   3240
         Width           =   5300
      End
      Begin VB.OptionButton optNegFmt 
         Caption         =   "Monetary symbol, space, number, negative sign"
         Height          =   255
         Index           =   11
         Left            =   240
         TabIndex        =   20
         Top             =   3000
         Width           =   5300
      End
      Begin VB.OptionButton optNegFmt 
         Caption         =   "Number, space, monetary symbol, negative sign"
         Height          =   255
         Index           =   10
         Left            =   240
         TabIndex        =   19
         Top             =   2760
         Width           =   5300
      End
      Begin VB.OptionButton optNegFmt 
         Caption         =   "Negative sign, monetary symbol, space, number"
         Height          =   255
         Index           =   9
         Left            =   240
         TabIndex        =   18
         Top             =   2520
         Width           =   5300
      End
      Begin VB.OptionButton optNegFmt 
         Caption         =   "Negative sign, number, space, monetary symbol"
         Height          =   255
         Index           =   8
         Left            =   240
         TabIndex        =   17
         Top             =   2280
         Width           =   5300
      End
      Begin VB.OptionButton optNegFmt 
         Caption         =   "Number, monetary symbol, negative sign"
         Height          =   255
         Index           =   7
         Left            =   240
         TabIndex        =   16
         Top             =   2040
         Width           =   5300
      End
      Begin VB.OptionButton optNegFmt 
         Caption         =   "Number, negative sign, monetary symbol"
         Height          =   255
         Index           =   6
         Left            =   240
         TabIndex        =   15
         Top             =   1800
         Width           =   5300
      End
      Begin VB.OptionButton optNegFmt 
         Caption         =   "Negative sign, number, monetary symbol"
         Height          =   255
         Index           =   5
         Left            =   240
         TabIndex        =   14
         Top             =   1560
         Width           =   5300
      End
      Begin VB.OptionButton optNegFmt 
         Caption         =   "Left parenthesis, number, monetary symbol, right parenthesis"
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   13
         Top             =   1320
         Width           =   5300
      End
      Begin VB.OptionButton optNegFmt 
         Caption         =   "Monetary symbol, number, negative sign"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   12
         Top             =   1080
         Width           =   5300
      End
      Begin VB.OptionButton optNegFmt 
         Caption         =   "Monetary symbol, negative sign, number"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   11
         Top             =   840
         Width           =   5300
      End
      Begin VB.OptionButton optNegFmt 
         Caption         =   "Negative sign, monetary symbol, number"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   10
         Top             =   600
         Width           =   5300
      End
      Begin VB.OptionButton optNegFmt 
         Caption         =   " Left parenthesis,monetary symbol,number,right parenthesis"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   9
         Top             =   360
         Width           =   5300
      End
   End
   Begin VB.Frame frPosFmt 
      Caption         =   "Positive value formats"
      Height          =   1455
      Left            =   240
      TabIndex        =   3
      Top             =   2160
      Width           =   2415
      Begin VB.OptionButton optPosFmt 
         Caption         =   "Suffix, separation"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   7
         Top             =   1080
         Width           =   1935
      End
      Begin VB.OptionButton optPosFmt 
         Caption         =   "Prefix, separation"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   6
         Top             =   840
         Width           =   1935
      End
      Begin VB.OptionButton optPosFmt 
         Caption         =   "Suffix, no separation"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   5
         Top             =   600
         Width           =   1935
      End
      Begin VB.OptionButton optPosFmt 
         Caption         =   "Prefix, no separation"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   1935
      End
   End
   Begin VB.Frame frCurrFormat 
      Caption         =   "Currency format"
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5535
      Begin VB.OptionButton optCurrFormat 
         Caption         =   "Use user settings"
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   2
         Top             =   600
         Width           =   2535
      End
      Begin VB.OptionButton optCurrFormat 
         Caption         =   "Use system settings"
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   2175
      End
   End
   Begin VB.Label Label7 
      Caption         =   "Cick the usercontrol and enter a value"
      Height          =   735
      Left            =   240
      TabIndex        =   38
      Top             =   1320
      Width           =   1335
   End
End
Attribute VB_Name = "frmDemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Demo Currency usercontrol
' For extended doc see code of usercontrol
' See also Form_Load event for initialising the usercontrol
'
' Paul Turcksin   June 2005


Option Explicit

Private Sub cmdApply_Click()
' apply the changes to currency symbol, number digits,...
      ucCurrency1.CurrSymbol = fCurrSymbol.Text
      ucCurrency1.CurrDigits = fDigits.Text
      ucCurrency1.DecimalSep = fDecSep.Text
      ucCurrency1.ThousandsSep = fThouSep.Text
      ucCurrency1.ThousandsGroup = fGroup
      ucCurrency1.CurrLeadingZero = chLeading.Value
End Sub

Private Sub cmdShow_Click()
   MsgBox ucCurrency1.Text
End Sub

Private Sub Form_Load()
' start with the system defaults
   optCurrFormat(0).Value = True
' uncomment line to put a value in the user control different
' from the default "0":
'   ucCurrency1.Text = "12345"
' or:
'   ucCurrency1.Text = 1234.5
End Sub

Private Sub optCurrFormat_Click(Index As Integer)
' toggle system defaults/user defined
   ucCurrency1.CurrFormat = Index
' de/activate frames
   frPosFmt.Visible = Index
   frNegFmt.Visible = Index
   frMisc.Visible = Index
' init options and variables in frames with system settings
   If Index Then
      optPosFmt(ucCurrency1.CurrPosFmt).Value = True
      optNegFmt(ucCurrency1.CurrNegFmt).Value = True
      fCurrSymbol.Text = ucCurrency1.CurrSymbol
      fDigits.Text = ucCurrency1.CurrDigits
      fDecSep.Text = ucCurrency1.DecimalSep
      fThouSep.Text = ucCurrency1.ThousandsSep
      fGroup = ucCurrency1.ThousandsGroup
      chLeading.Value = ucCurrency1.CurrLeadingZero
      Me.Height = 9000
   Else
      Me.Height = 2700
   End If
End Sub

Private Sub optNegFmt_Click(Index As Integer)
' negative value formatting
   ucCurrency1.CurrNegFmt = Index
End Sub

Private Sub optPosFmt_Click(Index As Integer)
' positive value formatting
   ucCurrency1.CurrPosFmt = Index
End Sub
