Attribute VB_Name = "mdlScanF"
'ScanF family commands in VB
'written by Phlip Bradbury <phlipping@yahoo.com>
'You may use this in your program as long as credit is
'given to me.

' ScanF calls InputBox() with its parameters
'SScanF takes the string parameter
'FScanF calls Line Input on a file
'All functions interpret the format as in C, and then
'return the number of parameters matching
'VScanF, VSScanF and VFScanF simulate the similar
'functions in C which take va_list as a parameter, these
'omit the ParamArray keyword, allowing you to pass an array
'of parameters, rather than listing the parameters.
'SScanF shows a common use of this

'Although there are six functions offerring similar functionality,
'Only VSScanF contains the actual code. The other 5 simply
'generate the input string and call VSScanF

'It handles all the escape sequences in C
'except \" and \' as they would not help in VB
'It also handles all the parameters except things like
'%ld, %lf, etc, as VB handles that sort of thing internally
'Format handles escape sequences
'  \a   Alert (Bel)
'  \b   Backspace
'  \f   Form Feed
'  \n   Newline (Line Feed)
'  \r   Carriage Return
'  \t   Horizontal Tab
'  \v   Verical Tab
'  \ddd Octal character
'  \xdd Hexadecimal character
'and parameters
'%[flags][width]formattype
'flags:
' * match type, but do not write to parameter
'format types handled:
'  %d, %i number
'  %o     octal number
'  %x, %X hexadecimal number
'  %f     floating point number
'  %c     single character (gives ASCII value)
'  %s     String
'so eg %d matches a number and sets its value to the next
'parameter, %*d will match it but ignore it
'%5d will read exactly 5 bytes and interperet it as a
'number.
'use \\ to type a backslash and either \% or %% to type a
'percent sign
'also %c returns the ascii value of the character,
'if you want to recieve a single-character string
'use Chr() or %1s
'finally, a space matches any number (0 or more) white
'spaces (space, tab, newline, etc), to match a space use
'\40 or \x20
Option Explicit
Option Compare Binary

'String SCAN Formatted
'equiv to sscanf() in C
Public Function SScanF(ByVal InputString As String, ByVal FormatString As String, ParamArray Parms() As Variant) As Integer
  Dim A() As Variant, I As Integer
  'create a copy of Parms because we can't pass a ParamArray to a function
  If UBound(Parms) >= LBound(Parms) Then
    ReDim A(LBound(Parms) To UBound(Parms))
    For I = LBound(Parms) To UBound(Parms)
      A(I) = Parms(I)
    Next
  Else
    ReDim A(0 To 0)
  End If
  SScanF = VSScanF(FormatString, FormatString, A)
  'copy values back
  If UBound(Parms) >= LBound(Parms) Then
    For I = LBound(Parms) To UBound(Parms)
      Parms(I) = A(I)
    Next
  End If
End Function

'SCAN Formatted
'equiv to scanf() in C
Public Function ScanF(ByVal Prompt As String, ByVal Title As String, ByVal Default As String, ByVal FormatString As String, ParamArray Parms() As Variant) As Integer
  Dim A() As Variant, I As Integer, S As String
  If Title = "" Then
    S = InputBox(Prompt, , Default)
  Else
    S = InputBox(Prompt, Title, Default)
  End If
  If S = "" Then
    ScanF = 0
    Exit Function
  End If
  'create a copy of Parms because we can't pass a ParamArray to a function
  If UBound(Parms) >= LBound(Parms) Then
    ReDim A(LBound(Parms) To UBound(Parms))
    For I = LBound(Parms) To UBound(Parms)
      A(I) = Parms(I)
    Next
  Else
    ReDim A(0 To 0)
  End If
  ScanF = VSScanF(S, FormatString, A)
  'copy values back
  If UBound(Parms) >= LBound(Parms) Then
    For I = LBound(Parms) To UBound(Parms)
      Parms(I) = A(I)
    Next
  End If
End Function

'File SCAN Formatted
'equiv to fscanf() in C
Public Function FScanF(ByVal FileNum As Integer, ByVal FormatString As String, ParamArray Parms() As Variant) As Integer
  Dim A() As Variant, I As Integer, S As String
  'create a copy of Parms because we can't pass a ParamArray to a function
  If UBound(Parms) >= LBound(Parms) Then
    ReDim A(LBound(Parms) To UBound(Parms))
    For I = LBound(Parms) To UBound(Parms)
      A(I) = Parms(I)
    Next
  Else
    ReDim A(0 To 0)
  End If
  Line Input #FileNum, S
  FScanF = VSScanF(S, FormatString, A)
  'copy values back
  If UBound(Parms) >= LBound(Parms) Then
    For I = LBound(Parms) To UBound(Parms)
      Parms(I) = A(I)
    Next
  End If
End Function

'Variable-argument SCAN Formatted
'equiv to vscanf() in C
Public Function VScanF(ByVal Prompt As String, ByVal Title As String, ByVal Default As String, ByVal FormatString As String, Parms() As Variant) As Integer
  Dim S As String
  If Title = "" Then
    S = InputBox(Prompt, , Default)
  Else
    S = InputBox(Prompt, Title, Default)
  End If
  If S = 0 Then
    VScanF = 0
    Exit Function
  End If
  VScanF = VSScanF(S, FormatString, Parms)
End Function

'Variable-argument File SCAN Formatted
'equiv to vfscanf() in C
Public Function VFScanF(ByVal FileNum As Integer, ByVal FormatString As String, ParamArray Parms() As Variant) As Integer
  Dim S As String
  Line Input #FileNum, S
  VFScanF = VSScanF(S, FormatString, Parms)
End Function

'Variable-argument String SCAN Formatted
'equiv to vsscanf() in C
'this is where the actual work is done
Public Function VSScanF(ByVal InputString As String, ByVal FormatString As String, Parms() As Variant) As Integer
  'general
  Dim Char As String
  Dim CharInput As String
  Dim CharMatch As String
  'escape
  Dim NumberBuffer As String
  'parameters
  Dim ParamUpTo As Integer
  Dim Flags As String
  Dim Width As String
  Dim StrToMatch As String
  'for calculating %e and %g
  Dim Mantissa As Double, Exponent As Long
  'for calculating %g
  Dim AddStrPercentF As String, AddStrPercentE As String
  ParamUpTo = LBound(Parms)
  CharInput = NextChar(InputString)
  Do While FormatString <> "" And CharInput & InputString <> ""
    Char = NextChar(FormatString)
    Select Case Char
      Case " "
        While IsWhiteSpace(CharInput)
          CharInput = NextChar(InputString)
        Wend
      Case "\"
        Char = NextChar(FormatString)
        Select Case Char
          Case "a" 'alert (bell)
            CharMatch = Chr(7)
          Case "b" 'backspace
            CharMatch = vbBack
          Case "f" 'formfeed
            CharMatch = vbFormFeed
          Case "n" 'newline (linefeed)
            CharMatch = vbLf
          Case "r" 'carriage return
            CharMatch = vbCr
          Case "t" 'horizontal tab
            CharMatch = vbTab
          Case "v" 'vertical tab
            CharMatch = vbVerticalTab
          Case "0" To "9" 'octal character
            NumberBuffer = Char
            While InStr("01234567", Left(FormatString, 1)) And Len(FormatString) > 0
              NumberBuffer = NumberBuffer & NextChar(FormatString)
            Wend
            CharMatch = Chr(Oct2Dec(NumberBuffer))
          Case "x" 'hexadecimal character
            NumberBuffer = ""
            While InStr("0123456789ABCDEFabcdef", Left(FormatString, 1)) And Len(FormatString) > 0
              NumberBuffer = NumberBuffer & NextChar(FormatString)
            Wend
            CharMatch = Chr(Hex2Dec(NumberBuffer))
          Case "\" 'backslash
            CharMatch = "\"
          Case "%" 'percent
            CharMatch = "%"
          Case Else 'unrecognised
            CharMatch = Char
            Debug.Print "WARNING: Unrecognised escape sequence: \" & Char
        End Select
        If CharInput <> CharMatch Then
          Exit Do
        End If
        CharInput = NextChar(InputString)
      Case "%"
        Char = NextChar(FormatString)
        If Char = "%" Then
          'match % sign
          If CharInput <> "%" Then
            Exit Do
          End If
          CharInput = NextChar(InputString)
        Else
          Flags = ""
          Width = ""
          While Char = "*"
            Flags = Flags & Char
            Char = NextChar(FormatString)
          Wend
          While IsNumeric(Char)
            Width = Width & Char
            Char = NextChar(FormatString)
          Wend
          If Width <> "" Then
            StrToMatch = NextChar(InputString, CInt(Width))
          Else
            StrToMatch = ""
            Select Case Char
              Case "d", "i" 'decimal integer
                If CharInput = "-" Then
                  StrToMatch = StrToMatch & CharInput
                  CharInput = NextChar(InputString)
                End If
                While IsNumeric(CharInput)
                  StrToMatch = StrToMatch & CharInput
                  CharInput = NextChar(InputString)
                Wend
              Case "o" 'octal integer
                While InStr("01234567", CharInput)
                  StrToMatch = StrToMatch & CharInput
                  CharInput = NextChar(InputString)
                Wend
              Case "x", "X" 'hexadecimal integer
                'allow for 0x###
                If Left(CharInput & InputString, 1) = "0x" Then
                  CharInput = NextChar(InputString)
                  CharInput = NextChar(InputString)
                End If
                While InStr("0123456789ABCDEFabcdef", CharInput)
                  StrToMatch = StrToMatch & CharInput
                  CharInput = NextChar(InputString)
                Wend
              Case "f" 'decimal floating-point
                If CharInput = "-" Then
                  StrToMatch = StrToMatch & CharInput
                  CharInput = NextChar(InputString)
                End If
                While IsNumeric(CharInput) Or CharInput = "."
                  StrToMatch = StrToMatch & CharInput
                  CharInput = NextChar(InputString)
                Wend
                If CharInput = "e" Then
                  StrToMatch = StrToMatch & CharInput
                  CharInput = NextChar(InputString)
                End If
                If CharInput = "+" Or CharInput = "-" Then
                  StrToMatch = StrToMatch & CharInput
                  CharInput = NextChar(InputString)
                End If
                While IsNumeric(CharInput)
                  StrToMatch = StrToMatch & CharInput
                  CharInput = NextChar(InputString)
                Wend
              Case "c" 'single character
                StrToMatch = StrToMatch & CharInput
                CharInput = NextChar(InputString)
              Case "s" 'string
                StrToMatch = CharInput & InputString
                CharInput = ""
                InputString = ""
              End Select
          End If
          If StrToMatch = "" Then
            Exit Do
          End If
          If InStr(Flags, "*") = 0 Then
            Select Case Char
              Case "d", "i" 'decimal integer
                Parms(ParamUpTo) = CLng(StrToMatch)
              Case "o" 'octal integer
                Parms(ParamUpTo) = Oct2Dec(StrToMatch)
              Case "x", "X" 'hexadecimal integer
                Parms(ParamUpTo) = Hex2Dec(StrToMatch)
              Case "f" 'decimal floating-point
                Parms(ParamUpTo) = CDbl(StrToMatch)
              Case "c" 'single character, passed ASCII value
                Parms(ParamUpTo) = CByte(Asc(StrToMatch))
              Case "s" 'string
                Parms(ParamUpTo) = StrToMatch
              Case Else
                Debug.Print "WARNING: unrecognised parameter sequence: %" & Flags & Width & Char
            End Select
            ParamUpTo = ParamUpTo + 1
          End If
        End If
      Case Else
        If CharInput <> Char Then
          Exit Do
        End If
        CharInput = NextChar(InputString)
    End Select
  Loop
  VSScanF = ParamUpTo - LBound(Parms)
End Function

'Various helper functions

'returns the first character from a buffer and removes
'it from the buffer
Private Function NextChar(ByRef Buffer As String, Optional ByVal NumChars As Integer = 1) As String
  NextChar = Mid(Buffer, 1, NumChars)
  Buffer = Mid(Buffer, NumChars + 1)
End Function

'convert octal to decimal
Private Function Oct2Dec(ByVal Octal As String) As Long
  Dim I As Integer
  I = 0
  While Octal <> ""
    I = I * 8 + Val(NextChar(Octal))
  Wend
  Oct2Dec = I
End Function

'convert hexadecimal to decimal
Private Function Hex2Dec(ByVal Hexadecimal As String) As Long
  Hex2Dec = CLng("&H" & Hexadecimal)
End Function

Private Function IsWhiteSpace(ByVal Char As String) As Boolean
  Select Case Char
    Case " ", vbTab, vbVerticalTab, vbCr, vbLf, vbCrLf, vbFormFeed, vbNewLine, vbNullChar
      IsWhiteSpace = True
    Case Else
      IsWhiteSpace = False
  End Select
End Function
