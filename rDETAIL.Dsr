VERSION 5.00
Begin {BD4B4E61-F7B8-11D0-964D-00A0C9273C2A} rDETAIL 
   ClientHeight    =   11415
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   20505
   OleObjectBlob   =   "rDETAIL.dsx":0000
End
Attribute VB_Name = "rDETAIL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Section10_Format(ByVal pFormattingInfo As Object)
If Field11.Value = "1" Then
    tKond.SetText "01-Sangat Baik"
ElseIf Field11.Value = "2" Then
    tKond.SetText "02-Baik"
ElseIf Field11.Value = "3" Then
    tKond.SetText "03-Sedang"
Else
    tKond.SetText "04-Jelek"
End If
If Field12.Value = "1" Then
    tKons.SetText "01-Baja"
ElseIf Field12.Value = "2" Then
    tKons.SetText "02-Beton"
ElseIf Field12.Value = "3" Then
    tKons.SetText "03-Batu Bata"
Else
    tKons.SetText "04-Kayu"
End If
End Sub

Private Sub Section5_Format(ByVal pFormattingInfo As Object)
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
