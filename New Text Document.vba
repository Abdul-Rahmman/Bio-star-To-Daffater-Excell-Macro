Sub Final_M()
'
'This Macro convert bio star sheet to Daffater sheet (31) Days 
'
'
Dim i As integer
Dim j As integer
Dim o As integer
Dim n As integer
Dim r As integer

j = 2
o = 14
n=1

'Copy users
For i = 1 To 55
Sheets("Bio").Select
Cells.item(j, "C").Copy Sheets("Dafater").Range("B" & o)
j = j + 31
o = o + 1
Next i 

j = 2
o = 5
r = 14

'Copy Attendence
For i = 1 To 100
Sheets("Bio").Select
Range("I" & j).Copy Sheets("Dafater").Cells(r,o)
j = j + 1
o = o + 2
If n > 31 Then
n = 1
r = r + 1
End If
Next i
'
End Sub