VERSION 5.00
Begin {BD4B4E61-F7B8-11D0-964D-00A0C9273C2A} rDHKP 
   ClientHeight    =   9765
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   23865
   OleObjectBlob   =   "rDHKP.dsx":0000
End
Attribute VB_Name = "rDHKP"
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

Sub C_OBJEK()
C_OBJ = "SELECT * FROM QOBJEKPAJAK WHERE NOPQ='" & Field6.Value & "' ORDER BY NOPQ ASC"
openDB (C_OBJ)
If rPajak.RecordCount > 0 Then rPajak.MoveFirst
Do While Not rPajak.EOF
    OPNama.SetText rPajak!Nm_wp
rPajak.MoveNext
Loop
End Sub

Private Sub Section10_Format(ByVal pFormattingInfo As Object)
''C_OBJ = "SELECT * FROM PEMBAYARAN_SPPT WHERE THN_PAJAK_SPPT= '" & C_TAHUN & "'"
''openDB (C_OBJ)
''If rPajak.RecordCount > 0 Then rPajak.MoveFirst
''Do While Not rPajak.EOF
''    If Field16.Value = rPajak!KD_PROPINSI & "." & rPajak!KD_DATI2 & "." & rPajak!KD_KECAMATAN & "." & rPajak!KD_KELURAHAN & "." & rPajak!KD_BLOK & "-" & rPajak!NO_URUT & "." & rPajak!KD_JNS_OP Then
''        dBayar.SetText rPajak!TGL_PEMBAYARAN_SPPT
''    End If
''rPajak.MoveNext
''Loop
'C_OBJ = "SELECT * FROM PEMBAYARAN_SPPT WHERE THN_PAJAK_SPPT='" & C_TAHUN & "' AND ((KD_PROPINSI + '.' + KD_DATI2 + '.' + KD_KECAMATAN + '.' + KD_KELURAHAN + '.' + KD_BLOK + '-' + NO_URUT + '.' + KD_JNS_OP)='" & Field6.Value & "')"
'openDB (C_OBJ)
'If rPajak.RecordCount > 0 Then rPajak.MoveFirst
'jum = 0
'If Not rPajak.EOF Then
'    xBayar.SetText Format(Field11.Value - rPajak!JML_SPPT_YG_DIBAYAR, "#,#0")
'    'cTanggal.SetText rPajak!TGL_PEMBAYARAN_SPPT
'Else
'    xBayar.SetText Format(Field11.Value, "#,#0")
'    'cTanggal.SetText " "
''    Field11.Value = Field11.Value * 0
'    jum = jum + xBayar.Text * 1
'End If
'cTanggal.SetText jum
If UCase(Trim(Field8.Value)) <> "PAKPAK BHARAT" Or UCase(Trim(Field8.Value)) = "-" Or (Trim(Field8.Value)) = "" Or IsNull(Field8.Value) = True Then 'Or UCase(Trim(Field91.Value)) <> "-" Then
    Field25.Suppress = True
    tDesa.Suppress = False
Else
    Field25.Suppress = False
    tDesa.Suppress = True
End If
End Sub

Private Sub Section13_Format(ByVal pFormattingInfo As Object)
'Text24.Suppress = True
xTGL2.SetText "Salak,_________________________"
If cekTampil = 1 Then
    xTGL1.Suppress = False
    xTGL2.Suppress = True
Else
    
    xTGL1.Suppress = True
    xTGL2.Suppress = False
End If
End Sub

Private Sub Section14_Format(ByVal pFormattingInfo As Object)

'If Round(Field20.Value * 1, 0) = 0 Then
    tBayar.SetText bilang(Field37.Value)
    If cekTampil = 0 Then
        Text3.Suppress = True
        yTgl.Suppress = False
    Else
        Text3.Suppress = False
        yTgl.Suppress = True
    End If
'    tBayar2.Suppress = True
'    Text8.Suppress = True
'    Text10.SetText "Pokok Ketetapan :"
'Else
'    tBayar.SetText bilang(Field14.Value * 1)
'    tBayar2.SetText bilang(Field1.Value)
'    If cekTampil = 0 Then
'        Text3.Suppress = True
'        yTgl.Suppress = False
'    Else
'        Text3.Suppress = False
'        yTgl.Suppress = True
'    End If
'    tBayar2.Suppress = False
'    Text8.Suppress = False
'    Text10.SetText "Pokok Ketetapan Lama:"
'End If

End Sub

Private Sub Section4_Format(ByVal pFormattingInfo As Object)

'MsgBox CrossTab2.ColumnGroups(0).Field.Value
'tObjek1.SetText CrossTab2.ColumnGroups(1).Field.Value
'tObjek2.SetText Field19.Value
'tObjek3.SetText Field21.Value
'tObjek4.SetText Field17.Value
End Sub

Private Sub Section6_Format(ByVal pFormattingInfo As Object)
'C_OBJ = "SELECT * FROM REF_BUKU "
'openDB (C_OBJ)
'If rPajak.RecordCount > 0 Then rPajak.MoveFirst
'Do While Not rPajak.EOF
'    If Field11.Value >= rPajak!NILAI_MIN_BUKU And Field11.Value <= rPajak!NILAI_MAX_BUKU Then
'        tJJ.SetText "DAFTAR HIMPUNAN KETETAPAN PAJAK DAN PEMBAYARAN BUKU " & rPajak!KD_BUKU
'    End If
'rPajak.MoveNext
'Loop

End Sub

