VERSION 5.00
Begin {BD4B4E61-F7B8-11D0-964D-00A0C9273C2A} rKelas1 
   ClientHeight    =   9765
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15150
   OleObjectBlob   =   "rKelas1.dsx":0000
End
Attribute VB_Name = "rKelas1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Section4_Format(ByVal pFormattingInfo As Object)
xTGL2.SetText "Salak,_________________________"
If cekTampil = 1 Then
    xTGL2.Suppress = True
    tTanggal.Suppress = False
    tTanggal.SetText "Salak, " & Format(Now, "dd mmmm yyyy")
Else
    xTGL2.Suppress = False
    tTanggal.Suppress = True
End If
End Sub