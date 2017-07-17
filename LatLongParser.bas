'   Copyright (C) 2017  Arild M Johannessen
'
'   This program is free software: you can redistribute it and/or modify
'   it under the terms of the GNU General Public License as published by
'   the Free Software Foundation, either version 3 of the License, or
'   (at your option) any later version.
'   This program is distributed in the hope that it will be useful,
'   but WITHOUT ANY WARRANTY; without even the implied warranty of
'   MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'   GNU General Public License for more details.
'   You should have received a copy of the GNU General Public License
'   along with this program.  If not, see <http://www.gnu.org/licenses/>. */


Option Explicit

Function ParseLat(inn As String) As Double
Dim CountIntegers As Integer
Dim Integers As String
Dim Decimals As Double
Dim Sign As Integer
Dim dd As Double
Dim mm As Double
Dim ss As Double

    FindNumbers inn, CountIntegers, Integers, Decimals, Sign
    Select Case CountIntegers
        Case 1: 'D
            dd = Val(Integers) + Decimals
        Case 2: 'DD
            dd = Val(Integers) + Decimals
        Case 3: 'DMM (consider if this should be implemented as DDD)
            dd = Val(Left(Integers, 1))
            mm = Val(Right(Integers, 2)) + Decimals
        Case 4: 'DDMM
            dd = Val(Left(Integers, 2))
            mm = Val(Right(Integers, 2)) + Decimals
        Case 5: 'DMMSS
            dd = Val(Left(Integers, 1))
            mm = Val(Mid(Integers, 2, 2))
            ss = Val(Right(Integers, 2)) + Decimals
        Case 6: 'DDMMSS
            dd = Val(Left(Integers, 2))
            mm = Val(Mid(Integers, 3, 2))
            ss = Val(Right(Integers, 2)) + Decimals
        Case Else:
            Err.Raise 13, , "Invalid number format"
    End Select
    
    ParseLat = Sign * (dd + mm / 60 + ss / 3600)
End Function

Function ParseLon(inn As String) As Double
Dim CountIntegers As Integer
Dim Integers As String
Dim Decimals As Double
Dim Sign As Integer
Dim dd As Double
Dim mm As Double
Dim ss As Double

    FindNumbers inn, CountIntegers, Integers, Decimals, Sign
    Select Case CountIntegers
        Case 1: 'D
            dd = Val(Integers) + Decimals
        Case 2: 'DD
            dd = Val(Integers) + Decimals
        Case 3: 'DDD
            dd = Val(Integers) + Decimals
        Case 4: 'DDMM
            dd = Val(Left(Integers, 2))
            mm = Val(Right(Integers, 2)) + Decimals
        Case 5: 'DDDMM
            dd = Val(Left(Integers, 3))
            mm = Val(Right(Integers, 2)) + Decimals
        Case 6: 'DDMMSS
            dd = Val(Left(Integers, 2))
            mm = Val(Mid(Integers, 3, 2))
            ss = Val(Right(Integers, 2)) + Decimals
        Case 7: 'DDDMMSS
            dd = Val(Left(Integers, 3))
            mm = Val(Mid(Integers, 4, 2))
            ss = Val(Right(Integers, 2)) + Decimals
        Case Else:
            Err.Raise 13, , "Invalid number format"
    End Select
    
    ParseLon = Sign * (dd + mm / 60 + ss / 3600)
End Function

Function FindNumbers(ByVal Text As String, ByRef CountIntegers As Integer, ByRef Integers As String, ByRef Decimals As Double, ByRef Sign As Integer)
Dim innShort As String
Dim countDecimals As Integer
Dim startNumerals As Integer
Dim endNumerals As Integer
Dim endDecimals As Integer
Dim decimalPoint As Integer
Dim n As Integer
Dim c As String
Dim hemisphere As Integer

    'Init variables
    CountIntegers = 0
    Integers = ""
    Decimals = 0
    Sign = 0
    
    innShort = Replace(Text, "º", "")
    innShort = Replace(innShort, "°", "")
    innShort = Replace(innShort, """", "")
    innShort = Replace(innShort, "'", "")
    innShort = Replace(innShort, " ", "")
    
    startNumerals = -1
    endNumerals = -1
    decimalPoint = -1
    
    
    '----------------------
    'Find numerals
    '----------------------
    For n = 1 To Len(innShort)
        c = Mid(innShort, n, 1)
        If decimalPoint = -1 And IsNumeric(c) Then
            If startNumerals = -1 Then startNumerals = n
            endNumerals = n
            CountIntegers = CountIntegers + 1
            GoTo ContinueFor
        End If
        
        If c = "," Or c = "." Then
            decimalPoint = n
            GoTo ContinueFor
        End If
        
        If decimalPoint <> -1 And IsNumeric(c) Then
            endDecimals = n
            countDecimals = countDecimals + 1
            GoTo ContinueFor
        End If
        
ContinueFor:
    Next n
    
    
    'End if no decimals after decimalpoint
    If countDecimals = 0 And decimalPoint <> -1 Then Err.Raise 13, , "No numbers after decimalpoint"
    
    'End if no numerals
    If startNumerals = -1 Then Err.Raise 13, , "No numbers found"
    
    
    '----------------------
    'Find Hemisphere
    '----------------------
    If startNumerals > 1 Then
        c = UCase(Mid(innShort, startNumerals - 1, 1))
        If c = "N" Or c = "S" Or c = "-" Or c = "E" Or c = "W" Then
        'Found hemisphere in front of numerals
            If c = "S" Or c = "-" Or c = "W" Then
                hemisphere = -1
            Else
                hemisphere = 1
            End If
        End If
    Else
        'Search for hemisphere after numerals
        Dim LastDigit As Integer
        LastDigit = Max(endNumerals, endDecimals)
        If LastDigit < Len(innShort) Then
            c = UCase(Mid(innShort, LastDigit + 1, 1))
            If c = "N" Or c = "S" Or c = "E" Or c = "W" Then
            'Found hemisphere after numerals
                If c = "S" Or c = "W" Then
                    hemisphere = -1
                Else
                    hemisphere = 1
                End If
            End If
        End If
    End If
    If hemisphere = 0 Then
        'Hemisphere not found
        'Northern or Western assumed
        hemisphere = 1
    End If
    Integers = Mid(innShort, startNumerals, CountIntegers)
    If countDecimals > 0 Then _
        Decimals = Val("0." & Mid(innShort, decimalPoint + 1, countDecimals))
    
    Sign = hemisphere
End Function

Function DDMM(ByVal inn As Double, ResultType As Integer, Decimals As Integer, lat As Boolean) As Variant
'***********************************
' Resulttype
' _x - string (ResultType <  10)
' 1x - double (ResultType >= 10

' x0 -  D (No leading zero)
' x1 - DD (   Leading zero)
' x2 - DDMM
' x3 - DD MM
' x4 - DDMMSS
' x5 - DD MM SS
'***********************************

Dim dd As Double
Dim mm As Double
Dim ss As Double

Dim Neg As Boolean
Dim Direction As String
Dim FormatString1 As String
Dim FormatString2 As String
    
Dim ReturnDouble As Boolean

    If ResultType >= 10 Then
        ReturnDouble = True
        ResultType = ResultType - 10
    Else
        ReturnDouble = False
    End If
    
    dd = 0
    mm = 0
    ss = 0
    
    Dim multiplier As Double
    multiplier = 1
    Neg = False
    If inn < 0 Then
        Neg = True
        multiplier = -1
        inn = -inn
    End If
        
    Select Case ResultType
        Case 0 To 1
            dd = Round(inn, Decimals)
        Case 2 To 3
            dd = Int(inn)
            mm = Round(60# * (inn - dd), Decimals)
        Case 4 To 5
            dd = Int(inn)
            mm = Int(60# * (inn - dd))
            ss = Round(3600# * (inn - dd) - 60# * mm, Decimals)
        Case Else
            DDMM = "Invalid Resulttype"
            Exit Function
    End Select
    
    'Roll up if rounding up to next whole unit
    If ss >= 60 Then
        ss = ss - 60
        mm = mm + 1
    End If
    If mm >= 60 Then
        mm = mm - 60
        dd = dd + 1
    End If
        
    If lat And Not Neg Then Direction = "N"
    If lat And Neg Then Direction = "S"
    If Not lat And Not Neg Then Direction = "E"
    If Not lat And Neg Then Direction = "W"
    


    If lat Then FormatString1 = "00" _
           Else FormatString1 = "000"
           
    


Select Case Decimals
        Case -1: FormatString2 = "00"
        Case 0: FormatString2 = "00"
        Case 1: FormatString2 = "00.0"
        Case 2: FormatString2 = "00.00"
        Case 3: FormatString2 = "00.000"
        Case 4: FormatString2 = "00.0000"
        Case 5: FormatString2 = "00.00000"
        Case 6: FormatString2 = "00.000000"
        Case Else: FormatString2 = "00.#"
    End Select
            
    Select Case ResultType
    Case 0: FormatString2 = Mid(FormatString2, 2, Len(FormatString2) - 1) 'Remove leading zero
    Case 1: If Not lat Then FormatString2 = "0" & FormatString2           'If longitude, include three zeros
    Case Else 'No Change in formatstring if Resulttype > 1
    End Select
    



    
    'Output the result as double or string
    If ReturnDouble Then
        'Return a double rounded to the precision given as a combination of decimalpoints and resulttype
        DDMM = multiplier * (dd + mm / 60# + ss / 3600#)
    Else
        'Return a string formatted as given in resulttype
        Select Case ResultType
            Case 0 To 1: DDMM = Format(dd, FormatString2) & Direction
            Case 2:      DDMM = Format(dd, FormatString1) & Format(mm, FormatString2) & Direction
            Case 3:      DDMM = Format(dd, FormatString1) & " " & Format(mm, FormatString2) & Direction
            Case 4:      DDMM = Format(dd, FormatString1) & Format(mm, "00") & Format(ss, FormatString2) & Direction
            Case 5:      DDMM = Format(dd, FormatString1) & " " & Format(mm, "00") & " " & Format(ss, FormatString2) & Direction
            Case Else
                DDMM = "Invalid Resulttype"
                Exit Function
        End Select
    End If
End Function

Function Max(i1 As Integer, i2 As Integer) As Integer
    If i1 > i2 Then Max = i1 Else Max = i2
End Function


