VERSION 5.00
Begin {BD4B4E61-F7B8-11D0-964D-00A0C9273C2A} rParkir 
   ClientHeight    =   9765
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15375
   OleObjectBlob   =   "rParkir.dsx":0000
End
Attribute VB_Name = "rParkir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Section9_Format(ByVal pFormattingInfo As Object)
Dim KLS(2), NILAI(2)
C_STR = "SELECT * FROM  DBKB_JPB6 WHERE THN_DBKB_JPB6 ='" & C_TAHUN & "'"
openDB (C_STR)
If rPajak.RecordCount > 0 Then rPajak.MoveFirst
i = 0
Do While Not rPajak.EOF
i = i + 1
    KLS(i) = rPajak!KLS_DBKB_JPB6
    NILAI(i) = rPajak!NILAI_DBKB_JPB6
rPajak.MoveNext
Loop
TKELAS.SetText KLS(1)
TKELAS1.SetText KLS(2)
tNilai.SetText Format(NILAI(1), "#,#0")
TNILAI1.SetText Format(NILAI(2), "#,#0")
End Sub
