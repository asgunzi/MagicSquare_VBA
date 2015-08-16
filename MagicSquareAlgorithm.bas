Attribute VB_Name = "MagicSquareAlgorithm"
Option Explicit

Sub magicsquare(ByVal size As Integer, ByRef Square As Variant)
'May-2010 / Aug 2015 by Arnaldo Gunzi
'Receives the size of the square
'returns the MagicSquare in outSquare
'algorithms from: mathworld.wolfram.com

Dim i As Long, j As Long
Dim posX As Long, posY As Long
Dim posXtest As Long, posYtest As Long
Dim tam2 As Long

Dim m As Long
Dim aux() As String
Dim aux2() As Integer



If size < 3 Then
    MsgBox "The square must have size >= 3"
    Exit Sub
End If

ReDim Square(1 To size, 1 To size)


If size Mod 2 = 1 Then
    'For odd size magic Squares
    posX = (size + 1) / 2
    posY = 1
    
    For i = 1 To size ^ 2
        Square(posY, posX) = i
        
            If posY = 1 Then
                posYtest = size
            Else
                posYtest = posY - 1
            End If
        
            If posX = 1 Then
                posXtest = size
            Else
                posXtest = posX - 1
            End If
        
        If Square(posYtest, posXtest) = 0 Then
            posY = posYtest
            posX = posXtest
        Else
            posY = posY + 1
        End If
    Next i

ElseIf size Mod 4 = 0 Then
    'size 4
    For i = 1 To size
        For j = 1 To size
            If i Mod 4 = 1 Or i Mod 4 = 0 Then
                If j Mod 4 = 0 Or j Mod 4 = 1 Then
                    Square(i, j) = size ^ 2 - (size * (i - 1) + j) + 1
                Else
                    Square(i, j) = size * (i - 1) + j
                End If
            Else
                If j Mod 4 = 2 Or j Mod 4 = 3 Then
                    Square(i, j) = size ^ 2 - (size * (i - 1) + j) + 1
                Else
                    Square(i, j) = size * (i - 1) + j
                End If
            End If
        Next j
    Next i
Else
    'Sie 6
    'LUX method
     m = (size - 2) / 4
     
     ReDim aux(1 To 2 * m + 1, 1 To 2 * m + 1)
     ReDim aux2(1 To 2 * m + 1, 1 To 2 * m + 1)

     For i = 1 To m + 1
        For j = 1 To 2 * m + 1
            aux(i, j) = "L"
        Next j
     Next i
     
    For j = 1 To 2 * m + 1
        aux(m + 2, j) = "U"
    Next j
    
     For i = m + 3 To 2 * m + 1
        For j = 1 To 2 * m + 1
            aux(i, j) = "X"
        Next j
     Next i
    
    
    aux(m + 1, m + 1) = "U"
    aux(m + 2, m + 1) = "L"
    
    tam2 = 2 * m + 1
    posX = (tam2 + 1) / 2
    posY = 1
    
    For i = 1 To tam2 ^ 2
        aux2(posY, posX) = i
        
        If aux(posY, posX) = "L" Then
            Square(2 * (posY - 1) + 1, 2 * (posX - 1) + 2) = 4 * (i - 1) + 1
            Square(2 * (posY - 1) + 2, 2 * (posX - 1) + 1) = 4 * (i - 1) + 2
            Square(2 * (posY - 1) + 2, 2 * (posX - 1) + 2) = 4 * (i - 1) + 3
            Square(2 * (posY - 1) + 1, 2 * (posX - 1) + 1) = 4 * (i - 1) + 4
            
        ElseIf aux(posY, posX) = "U" Then
            Square(2 * (posY - 1) + 1, 2 * (posX - 1) + 1) = 4 * (i - 1) + 1
            Square(2 * (posY - 1) + 2, 2 * (posX - 1) + 1) = 4 * (i - 1) + 2
            Square(2 * (posY - 1) + 2, 2 * (posX - 1) + 2) = 4 * (i - 1) + 3
            Square(2 * (posY - 1) + 1, 2 * (posX - 1) + 2) = 4 * (i - 1) + 4
        ElseIf aux(posY, posX) = "X" Then
            Square(2 * (posY - 1) + 1, 2 * (posX - 1) + 1) = 4 * (i - 1) + 1
            Square(2 * (posY - 1) + 2, 2 * (posX - 1) + 2) = 4 * (i - 1) + 2
            Square(2 * (posY - 1) + 2, 2 * (posX - 1) + 1) = 4 * (i - 1) + 3
            Square(2 * (posY - 1) + 1, 2 * (posX - 1) + 2) = 4 * (i - 1) + 4

        End If
            
            If posY = 1 Then
                posYtest = tam2
            Else
                posYtest = posY - 1
            End If
        
            If posX = tam2 Then
                posXtest = 1
            Else
                posXtest = posX + 1
            End If
        
        If aux2(posYtest, posXtest) = 0 Then
            posY = posYtest
            posX = posXtest
        Else
            posY = posY + 1
        End If
    Next i
    
End If

End Sub

