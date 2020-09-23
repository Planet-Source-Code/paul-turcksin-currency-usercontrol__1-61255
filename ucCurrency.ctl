VERSION 5.00
Begin VB.UserControl ucCurrency 
   ClientHeight    =   405
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   990
   ScaleHeight     =   405
   ScaleWidth      =   990
   ToolboxBitmap   =   "ucCurrency.ctx":0000
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   0
      Width           =   975
   End
End
Attribute VB_Name = "ucCurrency"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Inspired by PSC submission: Euro / Dollar input box without the use of Maskedit
'
' http://www.Planet-Source-Code.com/vb/default.asp?lngCId=61011&lngWId=1
'
' This usercontrol uses a standard textbox with one main property:
'
'    - CurrFomat: use local currency system settings (0 - System)(default)
'      or allow changes to system settings (1 - User).
'
' In the latter case additional properties are available.
'
'  - CurrSymbol: Currency Character (up to 6 characters)
'  - CurrDigits: Number of fractional digits (0 - 9)
'  - DecimalSep: decimal separator character (1 - 3 characters)
'  - ThousandsSep: thousands separator character (1 - 3 characters)
'  - ThousandsGroup: thousands grouping (funny it isn't always 3) (0 - 9)
'  - CurrPosFmt: Positive value formatting: currency symbol left/right of the numeric
'    value with or without 1 space separation (4 formats)
'  - CurrNegFmt: negative value formatting: negative sign or parenthesis, currency
'    symbol left/right of the numeric value, space separation; in various different
'    orders (16 formats)
'  - LeadinZero: (0- NoLeadingZero) or (1 - LeadingZero)
'
' Additional property:SelectAll: highlight value when it gets the focus
'                                (0 - Manual) user has to select
'                                (1 - OnFocus)
'
' Other features:
' - values can be entered without taking into account the formatting.
'   Example: the textbox contains "$ 123.50". You can overwrite this with
'            "5.5". When the usercontrol looses foces it will contain "$ 5.50".
' - formatting can be programmatically changed on the spot.
'
' If the user enters a faulty value the backcolor turns red and computer beeps
'
' The value returned to the user program is a string stripped of all non-numeric
' characters except a minus sign for negative values and (always) a decimal Dot.

' !!! As this usercontrol doesn't have a default property the text property
'     must always be mentioned in the program when setting or retreiving the
'     value of this control (Ex:  yVar = ucCurrency1.text)

' Paul Turcksin  June 2005
'
'___________________________________________________________________________________

Option Explicit
'Default Property Values:
Const m_def_CurrFormat = 0
'Property Variables:
Dim m_CurrDigits As Long
Dim m_CurrFormat As Integer
Dim m_CurrNegFmt As Long
Dim m_CurrPosFmt As Long
Dim m_CurrSymbol As String
Dim m_DecimalSep As String
Dim m_LeadingZero As Long
Dim m_ThousandsSep As String
Dim m_ThousandsGroup As Integer
Dim m_SelectAll As Integer
' preserve backcolor
Dim OriginalBackcolor As Long
' current value
Dim m_Value As String   ' stripped of symbol, thousands separator, space separator
                        ' positive sign, parenthesis turned into negative sign
                        ' and a decimal dot
' current currency format
Dim ThisCurrFmt As CURRENCYFMT

' Properties enumerations ----------------------------------------------------------
' TEXTBOX appearance
Public Enum AppearanceConstants
    Flat = 0
    is3D = 1
End Enum

' TEXTBOX borderline
Public Enum BorderlineConstants
   no = 0
   FixedSingle = 1
End Enum
   
' TEXTBOX SELECT ALL
Public Enum SelectAllConstants
   Manual = 0
   OnFocus = 1
End Enum

' CURRENCY format
Public Enum CurrencyFormatConstants
   System = 0
   User = 1
End Enum

' CURRENCY Leading zero
Public Enum CurrencyLeadingZeroConstants
   NoLeadingZero = 0
   LeadingZero = 1
End Enum

' negative currency formats
Public Enum CurrencyNegFormatConstants
   ParSymVal = 0
   SignSymVal = 1
   SymSignVal = 2
   SymValSign = 3
   ParValSym = 4
   SignValSym = 5
   ValSignSym = 6
   ValSymSign = 7
   SignValSepSym = 8
   SignSymSepVal = 9
   ValSepSymSign = 10
   SymSepValSign = 11
   SymSepSignVal = 12
   ValSignSepSym = 13
   ParSymSepVal = 14
   ParValSepSym = 15
End Enum

' positive currency formats
Public Enum CurrencyPosFormatConstants
   PrefixNoSep = 0
   SuffixNoSep = 1
   PrefixSep = 2
   SuffixSep = 3
End Enum

'Event Declarations:

' System currency settings stuff

Private Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long
Private Const LOCALE_USER_DEFAULT = &H400

'String used as the local monetary symbol.
' The maximum number of characters allowed for this string is six.
Private Const LOCALE_SCURRENCY = &H14        '  local monetary symbol

'Number of fractional digits for the local monetary format.
'The maximum number of characters allowed for this string is three.
Private Const LOCALE_ICURRDIGITS = &H19

' Character(s) used as the monetary decimal separator.
' The maximum number of characters allowed for this string is four.
Private Const LOCALE_SMONDecimalSep = &H16

' Character(s) used as the monetary separator between groups of digits to the left
' of the decimal.
' The maximum number of characters allowed for this string is four.
Private Const LOCALE_SMONTHOUSANDSEP = &H17

' Sizes for each group of monetary digits to the left of the decimal.
'An explicit size is needed for each group, and sizes are separated by semicolons.
' If the last value is zero, the preceding value is repeated.
' For example, to group thousands, specify 3;0. Indic languages group the first
' thousand and then group by hundreds—for example, 12,34,56,789, which is represented
' by 3;2;0. !!! how to this with only 3 char ???
' The maximum number of characters allowed for this string is four.
Private Const LOCALE_SMONGROUPING = &H18

' Specifier for leading zeros in decimal fields.
' The maximum number of characters allowed for this string is two.
' The specifier can be one of the following values.
' 0 No leading zeros
' 1 Leading zeros
Private Const LOCALE_ILZERO = &H12

' Negative currency mode.
' The maximum number of characters allowed for this string is three.
' The mode can be one of the following values.
' 0 Left parenthesis,monetary symbol,number,right parenthesis.Example: ($1.1)
' 1 Negative sign, monetary symbol, number. Example: -$1.1
' 2 Monetary symbol, negative sign, number. Example: $-1.1
' 3 Monetary symbol, number, negative sign. Example: $1.1-
' 4 Left parenthesis, number, monetary symbol, right parenthesis. Example: (1.1$)
' 5 Negative sign, number, monetary symbol. Example: -1.1$
' 6 Number, negative sign, monetary symbol. Example: 1.1-$
' 7 Number, monetary symbol, negative sign. Example: 1.1$-
' 8 Negative sign, number, space, monetary symbol (like #5, but with a space before the monetary symbol). Example: -1.1 $
' 9 Negative sign, monetary symbol, space, number (like #1, but with a space after the monetary symbol). Example: -$ 1.1
' 10 Number, space, monetary symbol, negative sign (like #7, but with a space before the monetary symbol). Example: 1.1 $-
' 11 Monetary symbol, space, number, negative sign (like #3, but with a space after the monetary symbol). Example: $ 1.1-
' 12 Monetary symbol, space, negative sign, number (like #2, but with a space after the monetary symbol). Example: $ -1.1
' 13 Number, negative sign, space, monetary symbol (like #6, but with a space before the monetary symbol). Example: 1.1- $
' 14 Left parenthesis, monetary symbol, space, number, right parenthesis (like #0, but with a space after the monetary symbol). Example: ($ 1.1)
' 15 Left parenthesis, number, space, monetary symbol, right parenthesis (like #4, but with a space before the monetary symbol). Example: (1.1 $)
Private Const LOCALE_INEGCURR = &H1C

' Position of the monetary symbol in the positive currency mode.
' The maximum number of characters allowed for this string is two.
' The mode can be one of the following values.
' 0 Prefix, no separation, for example $1.1
' 1 Suffix, no separation, for example 1.1$
' 2 Prefix, 1-character separation, for example $ 1.1
' 3 Suffix, 1-character separation, for example 1.1 $
Private Const LOCALE_ICURRENCY = &H1B


' The MSDN documentation is not always clear and subject to interpretation/guessing.
' This is especially true for grouping where it is contradictory. In theory one could
' define "variable" grouping Ex. 12,34,567.89 but this format cannot be retrieved because
' a maximum of 4 characters for the recieving buffer is allowed, just enough for fixes
' length grouping.
' For the number of fractional digits a constant exist for currency (LOCALE_ICURRDIGITS)
' but the doc specifies (LOCALE_IDIGITS) used for numeric values.
' The currency symbol LOCALE_SCURRENCY or LOCALE_SCURRENCY
' Constants between () as per MSDN doc
' Constants between [] not documented i.e. my choice
' Variables between {}
Private Type CURRENCYFMT
        NumDigits As Long          ' [LOCALE_ICURRDIGITS] number of fractional digits
        LeadingZero As Long        ' ( LOCALE_ILZERO) if leading zero in decimal fields Grouping
        Grouping As Long           ' {m-Group} group size left of decimal
        lpDecimalSep As String     ' [LOCALE_SMONDecimalSep] decimal separator string
        lpThousandSep As String    ' [LOCALE_SMONTHOUSANDSEP] thousand separator string
        NegativeOrder As Long      ' (LOCALE_INEGCURR) negative currency ordering
        PositiveOrder As Long      '(LOCALE_ICURRENCY) positive currency ordering
        lpCurrencySymbol As String ' [LOCALE_SCURRENCY] currency symbol string
End Type
Private Declare Function GetCurrencyFormat Lib "kernel32" Alias "GetCurrencyFormatA" (ByVal Locale As Long, ByVal dwFlags As Long, ByVal lpValue As String, lpFormat As CURRENCYFMT, ByVal lpCurrencyStr As String, ByVal cchCurrency As Long) As Long



'============================== PROPERTIES ==============================
'
' ................. Alignment
Public Property Get Alignment() As AlignmentConstants
   Alignment = Text1.Alignment
End Property
Public Property Let Alignment(ByVal New_Alignment As AlignmentConstants)
   Text1.Alignment() = New_Alignment
   PropertyChanged "Alignment"
End Property
' ................. Appearance
Public Property Get Appearance() As AppearanceConstants
   Appearance = Text1.Appearance
End Property
Public Property Let Appearance(ByVal New_Appearance As AppearanceConstants)
   Text1.Appearance() = New_Appearance
   PropertyChanged "Appearance"
End Property
' .................  Backcolor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
   BackColor = Text1.BackColor
End Property
Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
   Text1.BackColor() = New_BackColor
   PropertyChanged "BackColor"
End Property
' .................  Borderstyle
Public Property Get BorderStyle() As BorderlineConstants
   BorderStyle = Text1.BorderStyle
End Property
Public Property Let BorderStyle(ByVal New_BorderStyle As BorderlineConstants)
   Text1.BorderStyle() = New_BorderStyle
   PropertyChanged "BorderStyle"
End Property
' ................. Causes validation
Public Property Get CausesValidation() As Boolean
   CausesValidation = Text1.CausesValidation
End Property
Public Property Let CausesValidation(ByVal New_CausesValidation As Boolean)
   Text1.CausesValidation() = New_CausesValidation
   PropertyChanged "CausesValidation"
End Property
' ................. fractional digits
Public Property Get CurrDigits() As Integer
    CurrDigits = m_CurrDigits
End Property
Public Property Let CurrDigits(ByVal New_CurrDigits As Integer)
   If m_CurrFormat = System Then Exit Property
   If New_CurrDigits < 0 Or New_CurrDigits > 9 Then Exit Property
   m_CurrDigits = New_CurrDigits
   subUpdateCurrencyFormat "CurrDigits"
   PropertyChanged "CurrDigits"
  subShowIt
End Property
' ................. Currency format
Public Property Get CurrFormat() As CurrencyFormatConstants
   CurrFormat = m_CurrFormat
End Property
Public Property Let CurrFormat(ByVal New_CurrFormat As CurrencyFormatConstants)
   If New_CurrFormat < System Or New_CurrFormat > User Then Exit Property
   m_CurrFormat = New_CurrFormat
   subGetCurrencySystemSettings
   PropertyChanged "CurrFormat"
  subShowIt
End Property
' ................. Negative Currency format
Public Property Get CurrNegFmt() As CurrencyNegFormatConstants
   CurrNegFmt = m_CurrNegFmt
End Property
Public Property Let CurrNegFmt(ByVal New_CurrNegFmt As CurrencyNegFormatConstants)
   If m_CurrFormat = 0 Then Exit Property
   If New_CurrNegFmt < 0 Or New_CurrNegFmt > 15 Then Exit Property
   m_CurrNegFmt = New_CurrNegFmt
   subUpdateCurrencyFormat "CurrNegFmt"
   PropertyChanged "CurrNegFmt"
  subShowIt
End Property
' ................. Positive Currency format
Public Property Get CurrPosFmt() As CurrencyPosFormatConstants
   CurrPosFmt = m_CurrPosFmt
End Property
Public Property Let CurrPosFmt(ByVal New_CurrPosFmt As CurrencyPosFormatConstants)
   If m_CurrFormat = 0 Then Exit Property
   If New_CurrPosFmt < 0 Or New_CurrPosFmt > 4 Then Exit Property
   m_CurrPosFmt = New_CurrPosFmt
   subUpdateCurrencyFormat "CurrPosFmt"
   PropertyChanged "CurrPosFmt"
  subShowIt
End Property
' ................. Currency character
Public Property Get CurrSymbol() As String
   CurrSymbol = m_CurrSymbol
End Property
Public Property Let CurrSymbol(ByVal New_CurrSymbol As String)
   If m_CurrFormat = System Then Exit Property
   If Len(New_CurrSymbol) = 0 Or Len(New_CurrSymbol) > 6 Then
      MsgBox "In valid currency symbol", vbCritical Or vbOKOnly, UserControl.Name
      Exit Property
   End If
   m_CurrSymbol = New_CurrSymbol
   subUpdateCurrencyFormat "CurrSymbol"
   PropertyChanged "CurrSymbol"
  subShowIt
End Property
' .................  Enabled
Public Property Get Enabled() As Boolean
   Enabled = Text1.Enabled
End Property
Public Property Let Enabled(ByVal New_Enabled As Boolean)
   Text1.Enabled() = New_Enabled
   PropertyChanged "Enabled"
End Property
' ................. decimal separator character
Public Property Get DecimalSep() As String
   DecimalSep = m_DecimalSep
End Property
Public Property Let DecimalSep(ByVal New_DecimalSep As String)
   If m_CurrFormat = System Then Exit Property
   If Len(New_DecimalSep) = 0 Or Len(New_DecimalSep) > 3 Then Exit Property
   m_DecimalSep = New_DecimalSep
   subUpdateCurrencyFormat "DecimalSep"
   PropertyChanged "DecimalSep"
  subShowIt
End Property
' .................  Font
Public Property Get Font() As Font
   Set Font = Text1.Font
End Property
Public Property Set Font(ByVal New_Font As Font)
   Set Text1.Font = New_Font
   PropertyChanged "Font"
  subShowIt
End Property
' .................  Forecolor
Public Property Get ForeColor() As OLE_COLOR
   ForeColor = Text1.ForeColor
End Property
Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
   Text1.ForeColor() = New_ForeColor
   PropertyChanged "ForeColor"
End Property
' .................  LeadingZero
Public Property Get CurrLeadingZero() As CurrencyLeadingZeroConstants
   CurrLeadingZero = m_LeadingZero
End Property
Public Property Let CurrLeadingZero(ByVal New_LeadingZero As CurrencyLeadingZeroConstants)
   If m_CurrFormat = System Then Exit Property
   If New_LeadingZero < NoLeadingZero Or New_LeadingZero > LeadingZero Then Exit Property
   m_LeadingZero = New_LeadingZero
   subUpdateCurrencyFormat "LeadingZero"
   PropertyChanged "LeadingZero"
  subShowIt
End Property
' ................. MaxLength
Public Property Get MaxLength() As Long
Attribute MaxLength.VB_Description = "Returns/sets the maximum number of characters that can be entered in a control."
   MaxLength = Text1.MaxLength
End Property
Public Property Let MaxLength(ByVal New_MaxLength As Long)
   Text1.MaxLength() = New_MaxLength
   PropertyChanged "MaxLength"
End Property

' ................. Select all
Public Property Get SelectAll() As SelectAllConstants
   SelectAll = m_SelectAll
End Property
Public Property Let SelectAll(ByVal New_SelectAll As SelectAllConstants)
   If New_SelectAll < Manual Or New_SelectAll > OnFocus Then Exit Property
   m_SelectAll = New_SelectAll
   PropertyChanged "SelectAll"
End Property
' ................... Text
Public Property Get Text() As String
Attribute Text.VB_Description = "Returns/sets the text contained in the control."
   Text = m_Value
End Property
Public Property Let Text(ByVal New_Text As String)
   m_Value = New_Text
  subShowIt
End Property
' ................. thousands separator character
Public Property Get ThousandsSep() As String
   ThousandsSep = m_ThousandsSep
End Property
Public Property Let ThousandsSep(ByVal New_ThousandsSep As String)
   If m_CurrFormat = System Then Exit Property
   If Len(New_ThousandsSep) = 0 Or Len(New_ThousandsSep) > 3 Then Exit Property
   m_ThousandsSep = New_ThousandsSep
   subUpdateCurrencyFormat "ThousandsSep"
   PropertyChanged "ThousandsSep"
  subShowIt
End Property
' ................. thousands grouping
' !!! see important remark in subGetCurrencySystemSettings
Public Property Get ThousandsGroup() As Integer
   ThousandsGroup = m_ThousandsGroup
End Property
Public Property Let ThousandsGroup(ByVal New_ThousandsGroup As Integer)
   If m_CurrFormat = System Then Exit Property
   If New_ThousandsGroup < 0 Or New_ThousandsGroup > 9 Then Exit Property
   m_ThousandsGroup = New_ThousandsGroup
   subUpdateCurrencyFormat "ThousandsGroup"
   PropertyChanged "ThousandsGroup"
  subShowIt
End Property
' .................. Tooltip
Public Property Get ToolTipText() As String
Attribute ToolTipText.VB_Description = "Returns/sets the text displayed when the mouse is paused over the control."
   ToolTipText = Text1.ToolTipText
End Property
Public Property Let ToolTipText(ByVal New_ToolTipText As String)
   Text1.ToolTipText() = New_ToolTipText
   PropertyChanged "ToolTipText"
End Property

' ================================== EVENTS =========================================

Private Sub UserControl_EnterFocus()
' make sure backcolor is reset to original backcolor
   Text1.BackColor = OriginalBackcolor
' select all
   If m_SelectAll = OnFocus Then
      Text1.SelStart = 0
      Text1.SelLength = Len(Text1.Text)
   End If
End Sub

Private Sub UserControl_ExitFocus()
' validate the data part of text1 content,
' if valid keep and show
   Dim strOut As String
   Dim strChar As String * 1
   Dim i As Integer
   
   
' extract significant part in "strOut"
   For i = 1 To Len(Text1.Text)
      strChar = Mid$(Text1.Text, i, 1)
      Select Case strChar
      
         Case "0" To "9"   ' numeric characters
            strOut = strOut & strChar
            
         Case m_DecimalSep ' decimal separator
            strOut = strOut & "."  ' !! internally always kept as a dot
            
         Case "-", "("     ' negative sign
            strOut = "-" & strOut
      End Select
   Next i
   
' numeric? (actual validation)
   If IsNumeric(strOut) Then
      m_Value = strOut
   Else
      Text1.BackColor = vbRed
      Beep
   End If
   
subShowIt

End Sub


' ============================== USER CONTROL ==============================

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
   m_CurrFormat = m_def_CurrFormat
   subGetCurrencySystemSettings
   m_Value = "0"
   m_SelectAll = OnFocus
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

   m_CurrDigits = PropBag.ReadProperty("CurrDigits", 2)
   m_CurrFormat = PropBag.ReadProperty("CurrFormat", 0)
   m_CurrNegFmt = PropBag.ReadProperty("CurrNegFmt", 0)
   m_CurrPosFmt = PropBag.ReadProperty("CurrPosFmt", 0)
   m_CurrSymbol = PropBag.ReadProperty("CurrSymbol", "")
   m_DecimalSep = PropBag.ReadProperty("DecimalSep", ".")
   m_LeadingZero = PropBag.ReadProperty("LeadingZero", 0)
   m_SelectAll = PropBag.ReadProperty("SelectAll", 1)
   m_ThousandsSep = PropBag.ReadProperty("ThousandsSep", "")
   m_ThousandsGroup = PropBag.ReadProperty("ThousandsGroup", 0)
   Text1.Appearance = PropBag.ReadProperty("Appearance", 1)
   Text1.Alignment = PropBag.ReadProperty("Alignment", 0)
   Text1.BackColor = PropBag.ReadProperty("BackColor", &H80000005)
   Text1.BorderStyle = PropBag.ReadProperty("BorderStyle", 1)
   Text1.CausesValidation = PropBag.ReadProperty("CausesValidation", True)
   Text1.Enabled = PropBag.ReadProperty("Enabled", True)
   Text1.ForeColor = PropBag.ReadProperty("ForeColor", &H80000008)
   Set Text1.Font = PropBag.ReadProperty("Font", Ambient.Font)
   Text1.MaxLength = PropBag.ReadProperty("MaxLength", 0)
   m_Value = PropBag.ReadProperty("Text", "")
   Text1.ToolTipText = PropBag.ReadProperty("ToolTipText", "")
   
   m_Value = "0"
   OriginalBackcolor = Text1.BackColor
' put them all in the currency formatting structure
   With ThisCurrFmt
      .Grouping = m_ThousandsGroup
      .LeadingZero = m_LeadingZero
      .lpCurrencySymbol = m_CurrSymbol & Chr(0)
      .lpDecimalSep = m_DecimalSep & Chr(0)
      .lpThousandSep = m_ThousandsSep & Chr(0)
      .NegativeOrder = m_CurrNegFmt
      .NumDigits = m_CurrDigits
      .PositiveOrder = m_CurrPosFmt
   End With
End Sub

Private Sub UserControl_Resize()
   Text1.Width = UserControl.Width
   Text1.Height = UserControl.Height
  subShowIt
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

   Call PropBag.WriteProperty("Appearance", Text1.Appearance, 1)
   Call PropBag.WriteProperty("Alignment", Text1.Alignment, 0)
   Call PropBag.WriteProperty("BackColor", Text1.BackColor, &H80000005)
   Call PropBag.WriteProperty("BorderStyle", Text1.BorderStyle, 1)
   Call PropBag.WriteProperty("CausesValidation", Text1.CausesValidation, True)
   Call PropBag.WriteProperty("CurrDigits", m_CurrDigits, 2)
   Call PropBag.WriteProperty("CurrFormat", m_CurrFormat, 0)
   Call PropBag.WriteProperty("CurrNegFmt", m_CurrNegFmt, 0)
   Call PropBag.WriteProperty("CurrPosFmt", m_CurrPosFmt, 0)
   Call PropBag.WriteProperty("CurrSymbol", m_CurrSymbol, "")
   Call PropBag.WriteProperty("DecimalSep", m_DecimalSep, ".")
   Call PropBag.WriteProperty("Enabled", Text1.Enabled, True)
   Call PropBag.WriteProperty("ForeColor", Text1.ForeColor, &H80000008)
   Call PropBag.WriteProperty("Font", Text1.Font, Ambient.Font)
   Call PropBag.WriteProperty("LeadingZero", m_LeadingZero, 0)
   Call PropBag.WriteProperty("MaxLength", Text1.MaxLength, 0)
   Call PropBag.WriteProperty("SelectAll", m_SelectAll, 1)
   Call PropBag.WriteProperty("ThousandsSep", m_ThousandsSep, "")
   Call PropBag.WriteProperty("ThousandsGroup", m_ThousandsGroup, 0)
   Call PropBag.WriteProperty("Text", m_Value, "Text")
   Call PropBag.WriteProperty("ToolTipText", Text1.ToolTipText, "")
End Sub

' ========================== SUPPORTING CODE ========================

Private Sub subGetCurrencySystemSettings()
   Dim ws As String
   Dim l As Long
   
' currency character
   l = 6
   ws = String(l, " ")
   l = GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SCURRENCY, ws, l)
   m_CurrSymbol = Left$(ws, l - 1)
' number fractional digits
   l = 3
   ws = String(l, " ")
   l = GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_ICURRDIGITS, ws, l)
   m_CurrDigits = Left$(ws, l - 1)
' leading zero
   l = 2
   ws = String(l, " ")
   l = GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_ILZERO, ws, l)
   m_LeadingZero = Left$(ws, l - 1)
' decimal separator
   l = 4
   ws = String(l, " ")
   l = GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SMONDecimalSep, ws, l)
   m_DecimalSep = Left$(ws, l - 1)
' thousands separator
   l = 4
   ws = String(l, " ")
   l = GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SMONTHOUSANDSEP, ws, l)
   m_ThousandsSep = Left$(ws, l - 1)
' thousands grouping
' The documentation says this string can be "3;0" (grouping per 3), but
' "3;2;0" (group of 3, then groups of 2) is also possible. However the function
' only accepts a max of 4 characters terminating 0 included. ???
' >>> I only keep from this format the first digit, which is also the requirement in
'     the CURRENCYFMT structure. Fairly easy to understand ;)
   l = 4
   ws = String(l, " ")
   l = GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SMONGROUPING, ws, l)
   m_ThousandsGroup = Mid(ws, 1, 1)
   

   
 ' negative format
   l = 3
   ws = String(l, " ")
   l = GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_INEGCURR, ws, l)
   m_CurrNegFmt = Left$(ws, l - 1)
   
 ' positive format
   l = 3
   ws = String(l, " ")
   l = GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_ICURRENCY, ws, l)
   m_CurrPosFmt = Left$(ws, l - 1)
   
' put them all in the currency formatting structure
   With ThisCurrFmt
      .Grouping = m_ThousandsGroup
      .LeadingZero = m_LeadingZero
      .lpCurrencySymbol = m_CurrSymbol & Chr(0)
      .lpDecimalSep = m_DecimalSep & Chr(0)
      .lpThousandSep = m_ThousandsSep & Chr(0)
      .NegativeOrder = m_CurrNegFmt
      .NumDigits = m_CurrDigits
      .PositiveOrder = m_CurrPosFmt
   End With
End Sub

Private Sub subShowIt()
   Dim ws As String
   Dim l As Long
   
' get length of string to be returned (last parameter=0)
   l = GetCurrencyFormat(LOCALE_USER_DEFAULT, 0, m_Value, ThisCurrFmt, ws, 0)
' prepare receiving buffer
   ws = String(l, " ")
' get formatted string
   l = GetCurrencyFormat(LOCALE_USER_DEFAULT, 0, m_Value, ThisCurrFmt, ws, l)
' show formatted value (strip terminating 0)
   Text1.Text = Left$(ws, l - 1)
End Sub

Private Sub subUpdateCurrencyFormat(ByVal Index As String)
   With ThisCurrFmt
      Select Case Index
         Case "CurrDigits":     .NumDigits = m_CurrDigits
         Case "CurrNegFmt":     .NegativeOrder = m_CurrNegFmt
         Case "CurrPosFmt":     .PositiveOrder = m_CurrPosFmt
         Case "CurrSymbol":     .lpCurrencySymbol = m_CurrSymbol & Chr(0)
         Case "DecimalSep":     .lpDecimalSep = m_DecimalSep & Chr(0)
         Case "LeadingZero":    .LeadingZero = m_LeadingZero
         Case "ThousandsGroup": .Grouping = m_ThousandsGroup
         Case "ThousandsSep":   .lpThousandSep = m_ThousandsSep & Chr(0)
      End Select
   End With
   
End Sub
