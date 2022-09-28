Attribute VB_Name = "Module1"
Public dbPajak As New ADODB.Connection
Public rPajak As New ADODB.Recordset
Public dbNJOP As New ADODB.Connection
Public rBUMI As New ADODB.Recordset
Public xID, NO_FORM, byPass, BYPASS1, BYPASS2, BYPASS3, bypass4
Public cTrans, nTipe_K
Public xxNon, xxJPB, K_JPB
Public DAYA_LISTRIK, JUM_SPLIT, JUM_WINDOW
Public LUAS_HRINGAN, LUAS_HSEDANG, LUAS_HBERAT, LUAS_HPENUTUP
Public JUM_LAP_BETON1, JUM_LAP_BETON2, JUM_LAP_ASPAL1, JUM_LAP_ASPAL2, JUM_LAP_RUMPUT1, JUM_LAP_RUMPUT2, JUM_LAP_BETON11, JUM_LAP_BETON21, JUM_LAP_ASPAL11, JUM_LAP_ASPAL21, JUM_LAP_RUMPUT11, JUM_LAP_RUMPUT21
Public PANJANG_PAGAR, BAHAN_PAGAR1, BAHAN_PAGAR2, LEBAR_TANGGA1, LEBAR_TANGGA2, JUM_LIFT1, JUM_LIFT2, JUM_LIFT3, JUM_PABX, DALAM_SUMUR
Public BAKAR_H, BAKAR_S, BAKAR_F
Public JLIFT(100), Nil_AC_Central(10)
Public nSistem, Luas_Kolam, JUM_AC_CENTRAL, JUM_GENSET, Nil_Boiler_Ht, Nil_Boiler_Ap, nMezanin, nDUKUNG
Public SELISIH_LUAS_EDIT
Public C_KEC, C_KEL, C_TAHUN, c_NOP, c_Ganti
Public J_CETAK
Public xLoad, nCom
Public nKomputer
Public zSEM, rSEM
Public ck_Ulin, tck_ulin
Public cBentuk, cekTampil
Public ccMenu
Public zJalan
Public xxLanjut
Public CetakQ
''===============================================
''Koneksi Database Menggunakan SQL 2008
''===============================================
Public Sub openDB(SQLStr As String)
pass = "sql0134"
dbName = "dbPajak"
If dbPajak.State = adStateOpen Then dbPajak.Close
Set dbPajak = Nothing
Set rPajak = Nothing
'dbPajak.Open "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=dbPajak;Data Source=" & nCom & ""
dbPajak.Open "Provider=SQLOLEDB.1;User ID=PBBQ;Password=0134;Initial Catalog=dbPajak;Data Source=" & nCom & ""
'dbPajak.Open "Provider=SQLNCLI10;SERVER=GOEDHAMCORPS-PC;Database=DBPAJAK;DataTypeCompatibility=80;User Id=PBBQ;Password=0134;"
'dbPajak.Open "Provider=SQLNCLI10;SERVER=" & nCom & ";Database=DBPAJAK;DataTypeCompatibility=80;User Id=PBBQ;Password=0134;"
rPajak.Open SQLStr, dbPajak, 1, 2
End Sub


'===============================================
'Koneksi Database Menggunakan ORACLE
'===============================================

'Public Sub openDB(SQLStr As String)
'On Error GoTo errKoneksi
'Set dbPajak = New ADODB.Connection
'Set rPajak = New ADODB.Recordset
'dbPajak.CursorLocation = adUseClient
'dbPajak.Open "Driver={Microsoft ODBC for Oracle};Server=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=LOCALHOST)(PORT=1521))(CONNECT_DATA=(SID=PBB)));Uid=system;pwd=0134;"
''dbPajak.Open "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=dbPajak"
'rPajak.Open SQLStr, dbPajak, 1, 2
'Exit Sub
'errKoneksi:
'MsgBox "Koneksi Gagal", vbCritical, "Error"
''End
'End Sub

'===============================================
'Koneksi Database Menggunakan File Access : .mdb
'===============================================
'Public Sub openDB(SQLStr As String)
'If dbPajak.State = adStateOpen Then dbPajak.Close
'Set dbPajak = Nothing
'Set rPajak = Nothing
'SUMBER1 = App.Path & "\Dokumen\PBB-p2.mdb"
'pass = "empatspasienam"
'dbPajak.Open "Provider=Microsoft.Jet.OLEDB.4.0;Jet OLEDB:Database Password='" & pass & "';Data Source='" & SUMBER1 & "';Persist Security Info=False"
'  rPajak.Open SQLStr, dbPajak, 1, 2
'End Sub
Public Sub Pesan(JENIS, ME_CAPTION, ICO_TEXT, INFO_TEXT)
frmMessage.Show
If JENIS = 1 Then 'PESAN ERROR
    frmMessage.iMess.Picture = LoadPicture(App.Path & "\ICO_stop.ico")
ElseIf JENIS = 2 Then 'PESAN INFORMASI
    frmMessage.iMess.Picture = LoadPicture(App.Path & "\ICO_INFO.ICO")
ElseIf JENIS = 3 Then 'PESAN REPAIR DEFAULT
    frmMessage.iMess.Picture = LoadPicture(App.Path & "\ICO_REPAIR.GIF")
ElseIf JENIS = 4 Then 'PESAN KESALAHAN
    frmMessage.iMess.Picture = LoadPicture(App.Path & "\ICO_STOP.ICO")
Else 'PESAN PERTANYAAN
    frmMessage.iMess.Picture = LoadPicture(App.Path & "\ICO_TANYA.ICO")
End If
frmMessage.Caption = ME_CAPTION
frmMessage.Label2.Caption = ICO_TEXT
frmMessage.Label1.Caption = INFO_TEXT
frmMessage.Show
End Sub
