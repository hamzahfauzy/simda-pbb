VERSION 5.00
Begin {BD4B4E61-F7B8-11D0-964D-00A0C9273C2A} rSSPDt 
   ClientHeight    =   9765
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15375
   OleObjectBlob   =   "rSSPDt.dsx":0000
End
Attribute VB_Name = "rSSPDt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Function bilang(X As Currency)
Dim triliun As Currency
Dim milyar As Currency
Dim juta As Currency
Dim ribu As Currency
Dim satu As Currency
Dim sen As Currency
Dim baca As String
'jika x = 0, maka dibaca 0
If X = 0 Then
    baca = angka(0, 1)
Else
'pisah masing-masing bagian untuk triliun, milyar, jutaan, ribu, rupiah dan sen
triliun = Int(X * 0.001 ^ 4)
milyar = Int((X - triliun * 1000 ^ 4) * 0.001 ^ 3)
juta = Int((X - triliun * 1000 ^ 4 - milyar * 1000 ^ 3) / 1000 ^ 2)
ribu = Int((X - triliun * 1000 ^ 4 - milyar * 1000 ^ 3 - juta * 1000 ^ 2) / 1000)
satu = Int(X - triliun * 1000 ^ 4 - milyar * 1000 ^ 3 - juta * 1000 ^ 2 - ribu * 1000)
sen = Int((X - Int(X)) * 1000)
'
If triliun > 0 Then
    baca = ratus(triliun, 5) + "Triliun "
End If
'
If milyar > 0 Then
    baca = baca + ratus(milyar, 4) + "Milyar "
End If
'
If juta > 0 Then
    baca = baca + ratus(juta, 3) + "Juta "
End If
'
If ribu > 0 Then
    baca = baca + ratus(ribu, 2) + "Ribu "
End If
'
If satu > 0 Then
    baca = baca + ratus(satu, 1) + "Rupiah "
Else
    baca = baca + "Rupiah "
End If
'
If sen > 0 Then
    baca = baca + ratus(sen, 0) + "Sen "
End If
End If
bilang = Left(baca, 1) & Mid(baca, 2)
End Function
Function ratus(X As Currency, posisi As Integer) As String
Dim a100 As Integer, a10 As Integer, a1 As Integer
Dim baca As String
a100 = Int(X * 0.01)
a10 = Int((X - a100 * 100) * 0.1)
a1 = Int(X - a100 * 100 - a10 * 10)
'
If a100 = 1 Then
    baca = "Seratus "
Else
    If a100 > 0 Then
        baca = angka(a100, 2) + "Ratus "
    End If
End If
'
If a10 = 1 Then
    baca = baca + angka(a10 * 10 + a1, 2)
Else
    If a10 > 0 Then
        baca = baca + angka(a10, 2) + "Puluh "
    End If
    If a1 > 0 Then
        If posisi = 2 And a100 = 0 And a10 = 0 Then
            baca = baca + angka(a1, 1)
        Else
            baca = baca + angka(a1, 2)
        End If
    End If
End If
ratus = baca
End Function

Function angka(X As Integer, posisi As Integer)
Select Case X
    Case 0: angka = "Nol"
    Case 1:
            If posisi = 2 Then
                angka = "Satu "
            Else
                angka = "Se"
            End If
    Case 2: angka = "Dua "
    Case 3: angka = "Tiga "
    Case 4: angka = "Empat "
    Case 5: angka = "Lima "
    Case 6: angka = "Enam "
    Case 7: angka = "Tujuh "
    Case 8: angka = "Delapan "
    Case 9: angka = "Sembilan "
    Case 10: angka = "Sepuluh "
    Case 11: angka = "Sebelas "
    Case 12: angka = "Dua Belas "
    Case 13: angka = "Tiga Belas "
    Case 14: angka = "Empat Belas "
    Case 15: angka = "Lima Belas "
    Case 16: angka = "Enam Belas "
    Case 17: angka = "Tujuh Belas "
    Case 18: angka = "Delapan Belas "
    Case 19: angka = "Sembilan Belas "
End Select
End Function

Private Sub Section6_Format(ByVal pFormattingInfo As Object)
'tBilang.SetText bilang(Field18.Value)
End Sub
