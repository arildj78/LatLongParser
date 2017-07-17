#--------------------------------------------------------------------------
#   LatLongParser v0.2.1
#--------------------------------------------------------------------------
#   Copyright (C) 2017  Arild M Johannessen
#
#   This program is free software: you can redistribute it and/or modify
#   it under the terms of the GNU General Public License as published by
#   the Free Software Foundation, either version 3 of the License, or
#   (at your option) any later version.
#   This program is distributed in the hope that it will be useful,
#   but WITHOUT ANY WARRANTY; without even the implied warranty of
#   MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
#   GNU General Public License for more details.
#   You should have received a copy of the GNU General Public License
#   along with this program.  If not, see <http://www.gnu.org/licenses/>. */
#--------------------------------------------------------------------------



def ParseLat(inn):
#Dim CountIntegers As Integer
#Dim Integers As String
#Dim Decimals As Double
#Dim Sign As Integer
    dd = 0.0
    mm = 0.0
    ss = 0.0

    result = FindNumbers(inn)
    CountIntegers = result[0]
    Integers = result[1]
    Decimals = result[2]
    Sign = result[3]
    LatLon = result[4]


    
if CountIntegers == 1:       #D
        dd = int(Integers) + Decimals
    elif CountIntegers == 2: #DD
        dd = int(Integers) + Decimals
    elif CountIntegers == 3: #DMM (consider if this should be implemented as DDD)
        dd = int(Integers[0])
        mm = int(Integers[-2:]) + Decimals
    elif CountIntegers == 4: #DDMM
        dd = int(Integers[:2])
        mm = int(Integers[-2:]) + Decimals
    elif CountIntegers == 5: #DMMSS
        dd = int(Integers[0])
        mm = int(Integers[1:3])
        ss = int(Integers[-2:]) + Decimals
    elif CountIntegers == 6: #DDMMSS
        dd = int(Integers[:2])
        mm = int(Integers[2:4])
        ss = int(Integers[-2:]) + Decimals
    else:
        raise ValueError("Invalid number format")    


    result = Sign * (dd + mm / 60 + ss / 3600)
    return result

def ParseLon(inn):
#Dim CountIntegers As Integer
#Dim Integers As String
#Dim Decimals As Double
#Dim Sign As Integer
    dd = 0.0
    mm = 0.0
    ss = 0.0

    result = FindNumbers(inn)
    CountIntegers = result[0]
    Integers = result[1]
    Decimals = result[2]
    Sign = result[3]
    LatLon = result[4]

    if CountIntegers == 1:   #D
        dd = int(Integers) + Decimals
    elif CountIntegers == 2: #DD
        dd = int(Integers) + Decimals
    elif CountIntegers == 3: #DDD
        dd = int(Integers) + Decimals
    elif CountIntegers == 4: #DDMM
        dd = int(Integers[:2])
        mm = int(Integers[-2:]) + Decimals
    elif CountIntegers == 5: #DDDMM
        dd = int(Integers[:3])
        mm = int(Integers[-2:]) + Decimals
    elif CountIntegers == 6: #DDMMSS
        dd = int(Integers[:2])
        mm = int(Integers[2:4])
        ss = int(Integers[-2:]) + Decimals
    elif CountIntegers == 7: #DDDMMSS
        dd = int(Integers[:3])
        mm = int(Integers[3:5])
        ss = int(Integers[-2:]) + Decimals
    else:
        raise ValueError("Invalid number format")    


    result = Sign * (dd + mm / 60 + ss / 3600)
    return result

def FindNumbers(Text):  # Return parameters are [CountIntegers, Integers, Decimals, Sign, LatLon]
#Dim innShort As String
#Dim startNumerals As Integer
#Dim endNumerals As Integer
#Dim endDecimals As Integer
#Dim decimalPoint As Integer
#Dim n As Integer
#Dim c As String
#Dim countDecimals As Integer
#Dim hemisphere As Integer

    #Init variables
    countDecimals = 0
    hemisphere = 0
    CountIntegers = 0
    Integers = ""
    Decimals = 0.0
    Sign = 0

    innShort = Text.replace(" ", "")        #0x20
    innShort = innShort.replace("º", "")    #0xBA
    innShort = innShort.replace("°", "")    #0xB0
    innShort = innShort.replace("""", "")   #0x22
    innShort = innShort.replace("'", "")    #0x27


    startNumerals = -1
    endNumerals = -1
    decimalPoint = -1
    endDecimals = -1

    #----------------------
    #Find numerals
    #----------------------
    for n in range(0, len(innShort)):
        c = innShort[n]
        if decimalPoint == -1 and c.isdigit():
            if startNumerals == -1: startNumerals = n
            endNumerals = n
            CountIntegers = CountIntegers + 1
            continue


        if c == "," or c == "." :
            decimalPoint = n
            continue


        if (decimalPoint != -1) and c.isdigit():
            endDecimals = n
            countDecimals = countDecimals + 1
            continue





#End if no decimals after decimalpoint
    if countDecimals == 0 and decimalPoint != -1 : raise ValueError("No numbers after decimalpoint")
    
    #End if no numerals
    if startNumerals == -1 :raise ValueError("No numbers found")


    #----------------------
    #Find Hemisphere
    #----------------------
    if startNumerals > 0:
        c = innShort[startNumerals - 1].upper()

        if c == "N" or c == "S" or c == "-" or c == "E" or c == "W" :
        #Found hemisphere in front of numerals
            if c == "S" or c == "-" or c == "W" :
                hemisphere = -1
            else:
                hemisphere = 1


            if c == "N" or c == "S":
                LatLon = "Lat"
            elif c == "E" or c == "W":
                LatLon = "Lon"

            
        
    else:
        #Search for hemisphere after numerals
        #Dim LastDigit As Integer
        LastDigit = max(endNumerals, endDecimals)
        if LastDigit < len(innShort)-1 :
            c = innShort[LastDigit + 1].upper()
                
            if c == "N" or c == "S" or c == "E" or c == "W" :
            #Found hemisphere after numerals
                if c == "S" or c == "W" :
                    hemisphere = -1
                else:
                    hemisphere = 1


                if c == "N" or c == "S":
                    LatLon = "Lat"
                elif c == "E" or c == "W":
                    LatLon = "Lon"






    if hemisphere == 0 :
        #Hemisphere not found
        #Northern or Western assumed
        hemisphere = 1
        LatLon = "Unk"
    
    Integers = innShort[startNumerals : startNumerals + CountIntegers]
    if countDecimals > 0 : 
        Decimals = float("0." + innShort[decimalPoint + 1 : decimalPoint + 1 + countDecimals])
    
    Sign = hemisphere
    return [CountIntegers, Integers, Decimals, Sign, LatLon]
    


def DDMM(inn, ResultType, Decimals, lat):
#***********************************
# Resulttype
# _x - string (ResultType <  10)
# 1x - double (ResultType >= 10

# x0 -  D (No leading zero)
# x1 - DD (   Leading zero)
# x2 - DDMM
# x3 - DD MM
# x4 - DDMMSS
# x5 - DD MM SS
#***********************************

#Dim dd As Double
#Dim mm As Double
#Dim ss As Double

#Dim Neg As Boolean
#Dim Direction As String
#Dim formatString1 As String
#Dim formatString2 As String
    
#Dim ReturnDouble As Boolean

    if ResultType >= 10:
        ReturnDouble = True
        ResultType = ResultType - 10
    else:
        ReturnDouble = False
    

    dd = 0
    mm = 0
    ss = 0
    
    #Dim multiplier As Double
    multiplier = 1
    Neg = False
    if inn < 0:
        Neg = True
        multiplier = -1
        inn = -inn
    

    
    if ResultType == 0 or ResultType == 1:
        dd = round(inn, Decimals)
    elif  ResultType == 2 or ResultType == 3:
        dd = int(inn)
        mm = round(60.0 * (inn - dd), Decimals)
    elif  ResultType == 4 or ResultType == 5:
        dd = int(inn)
        mm = int(60.0 * (inn - dd))
        ss = round(3600.0 * (inn - dd) - 60.0 * mm, Decimals)
    else:
        result = "Invalid Resulttype"
        raise ValueError("The value assigned to ResultType is not supported")    


    #Roll up if rounding up to next whole unit
    if ss >= 60:
        ss = ss - 60
        mm = mm + 1
    
    if mm >= 60:
        mm = mm - 60
        dd = dd + 1

    
    if       lat and not Neg: Direction = "N"
    elif     lat and     Neg: Direction = "S"
    elif not lat and not Neg: Direction = "E"
    elif not lat and     Neg: Direction = "W"

    if lat: formatString1 = "02.0f" #00
    else  : formatString1 = "03.0f" #000
    

    IntDigits=2                                      #00    Initialize variable and set result as two digits, padded with zero.
    if               ResultType == 0 : IntDigits = 1 #0     Decimal degrees (lat or long) with leading zeros disabled.
    elif not lat and ResultType == 1 : IntDigits = 3 #000   If decimal longitude, use three digits

    if Decimals > 0: formatString2 = "0{0}.{1}f".format(IntDigits + 1 + Decimals, Decimals)  #"Create string to format result"
    else           : formatString2 = "0{0}.{1}f".format(IntDigits               , 0       )  #"Create string to format result"






    






    #Output the result as double or string
    if ReturnDouble:
        #Return a double rounded to the precision given as a combination of decimalpoints and resulttype
        result = multiplier * (dd + mm / 60.0 + ss / 3600.0)
    else:
        #Return a string formatted as given in resulttype
        
        if   ResultType == 0                   : result = format(dd, formatString2) + Direction                                                                     #0 - D     #1 - DD
        elif ResultType == 1                   : result = format(dd, formatString2) + Direction                                                                     #0 - D     #1 - DD
        elif ResultType == 2                   : result = format(dd, formatString1) +       format(mm, formatString2) +                                   Direction #2 - DDMM
        elif ResultType == 3                   : result = format(dd, formatString1) + " " + format(mm, formatString2) +                                   Direction #3 - DD MM
        elif ResultType == 4                   : result = format(dd, formatString1) +       format(mm, "02.0f")       +       format(ss, formatString2) + Direction #4 - DDMMSS
        elif ResultType == 5                   : result = format(dd, formatString1) + " " + format(mm, "02.0f")       + " " + format(ss, formatString2) + Direction #5 - DD MM SS
        else:
            result = "Invalid Resulttype"
            raise ValueError("The value assigned to ResultType is not supported")    


    return result






#-----------------------------------
# LatLongParser END
#-----------------------------------