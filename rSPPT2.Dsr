VERSION 5.00
Begin {BD4B4E61-F7B8-11D0-964D-00A0C9273C2A} rSPPT2 
   ClientHeight    =   9765
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   14985
   OleObjectBlob   =   "rSPPT2.dsx":0000
End
Attribute VB_Name = "rSPPT2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Function bilang(x As Currency)
Dim triliun As Currency
Dim milyar As Currency
Dim juta As Currency
Dim ribu As Currency
Dim satu As Currency
Dim sen As Currency
Dim baca As String
'jika x = 0, maka dibaca 0
If x = 0 Then
    baca = angka(0, 1)
Else
'pisah masing-masing bagian untuk triliun, milyar, jutaan, ribu, rupiah dan sen
triliun = Int(x * 0.001 ^ 4)
milyar = Int((x - triliun * 1000 ^ 4) * 0.001 ^ 3)
juta = Int((x - triliun * 1000 ^ 4 - milyar * 1000 ^ 3) / 1000 ^ 2)
ribu = Int((x - triliun * 1000 ^ 4 - milyar * 1000 ^ 3 - juta * 1000 ^ 2) / 1000)
satu = Int(x - triliun * 1000 ^ 4 - milyar * 1000 ^ 3 - juta * 1000 ^ 2 - ribu * 1000)
sen = Int((x - Int(x)) * 1000)
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
Function ratus(x As Currency, posisi As Integer) As String
Dim a100 As Integer, a10 As Integer, a1 As Integer
Dim baca As String
a100 = Int(x * 0.01)
a10 = Int((x - a100 * 100) * 0.1)
a1 = Int(x - a100 * 100 - a10 * 10)
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

Function angka(x As Integer, posisi As Integer)
Select Case x
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

Private Sub Section10_Format(ByVal pFormattingInfo As Object)
tBayar.SetText bilang(Field11.Value)
xTGL2.SetText "Salak,_________________________"
If cekTampil = 1 Then
    xTGL1.Suppress = False
    xTGL2.Suppress = True
Else
    
    xTGL1.Suppress = True
    xTGL2.Suppress = False
End If

End Sub
Sub C_OBJEK()
C_OBJ = "SELECT * FROM QOBJEKPAJAK WHERE NOPQ='" & Field6.Value & "' ORDER BY NOPQ ASC"
openDB (C_OBJ)
If rPajak.RecordCount > 0 Then rPajak.MoveFirst
Do While Not rPajak.EOF
    OPNama.SetText rPajak!Nm_wp
rPajak.MoveNext
Loop
End Sub

Private Sub Section6_Format(ByVal pFormattingInfo As Object)
'MsgBox Field91.Field
If UCase(Trim(Field91.Value)) <> "PAKPAK BHARAT" Or UCase(Trim(Field91.Value)) = "-" Or (Trim(Field91.Value)) = "" Or IsNull(Field91.Value) = True Then 'Or UCase(Trim(Field91.Value)) <> "-" Then
    Field25.Suppress = True
    tDesa.Suppress = False
Else
    Field25.Suppress = False
'    tDesa.Suppress = True
End If
End Sub

