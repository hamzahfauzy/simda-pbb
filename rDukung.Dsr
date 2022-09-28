VERSION 5.00
Begin {BD4B4E61-F7B8-11D0-964D-00A0C9273C2A} rDukung 
   ClientHeight    =   9765
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15375
   OleObjectBlob   =   "rDukung.dsx":0000
End
Attribute VB_Name = "rDukung"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Section9_Format(ByVal pFormattingInfo As Object)
C_STR = "select * from DBKB_MEZANIN where THN_DBKB_MEZANIN='" & C_TAHUN & "'"
openDB (C_STR)
If rPajak.RecordCount > 0 Then rPajak.MoveFirst
Do While Not rPajak.EOF
 tKons.SetText rPajak!NILAI_DBKB_MEZANIN
rPajak.MoveNext
Loop
C_STR = "select * from DBKB_jpb14 where THN_DBKB_JPB14='" & C_TAHUN & "'"
openDB (C_STR)
If rPajak.RecordCount > 0 Then rPajak.MoveFirst
Do While Not rPajak.EOF
 tKanopi.SetText rPajak!NILAI_DBKB_JPB14
rPajak.MoveNext
Loop
End Sub
