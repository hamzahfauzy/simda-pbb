VERSION 5.00
Begin VB.Form frmMessage 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Disconnect!"
   ClientHeight    =   1425
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5130
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1425
   ScaleWidth      =   5130
   Begin VB.CommandButton cOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   5490
      TabIndex        =   4
      Top             =   900
      Width           =   660
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   1530
      Left            =   -45
      TabIndex        =   2
      Top             =   -120
      Width           =   1515
      Begin VB.Image iMess 
         Height          =   600
         Left            =   330
         Picture         =   "frmMessage.frx":0000
         Stretch         =   -1  'True
         Top             =   255
         Width           =   600
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Cek Dabatabse ADMIN di SERVER!"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   435
         Left            =   60
         TabIndex        =   3
         Top             =   930
         Width           =   1395
      End
   End
   Begin VB.Label LOK 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "OK!"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   300
      Left            =   2895
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   975
      Width           =   540
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "User tidak terhubung ke SERVER!"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   435
      Left            =   1665
      TabIndex        =   0
      Top             =   270
      Width           =   3330
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H00000000&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   345
      Left            =   2865
      Shape           =   4  'Rounded Rectangle
      Top             =   930
      Width           =   630
   End
End
Attribute VB_Name = "frmMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cOK_Click()
'LOK_Click
Unload Me
End Sub

Private Sub cOK_GotFocus()

Shape5.FillColor = vbBlue
End Sub

Private Sub Form_Activate()
Me.Top = (frmUtama.ScaleHeight - Me.Height) / 2
Me.Left = (frmUtama.Width - Me.Width) / 2
'MsgBox PesanQ
'If PesanQ = 1 Then
   ' Label1.Caption = "DATABASE tidak ditemukan...! " & _
                    vbCrLf & "RECONNECT to Database SERVER ?"
'    LOK.left = 1740
'    Shape5.left = 1725
'    LCancel.Visible = False
'    Shape1.Visible = False
'    iMess.Picture = LoadPicture(App.Path & "\repair.gif")
  '  Label2.Caption = "Cek Database SERVER!"
'ElseIf PesanQ = 2 Then
'    iMess.Picture = LoadPicture(App.Path & "\repair.gif")
 '   Label1.Caption = "RECONNECT to Database SERVER ?"
  '  Label2.Caption = "Cek Dabatabse utama di SERVER!"
'    LCancel.left = 1230
'    Shape1.left = 1185
'    LOK.left = 2235
'    Shape5.left = 2190
'    LCancel.Visible = True
'    Shape1.Visible = True
'End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Shape5.FillColor = &HC0C0C0
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub


Private Sub LOK_Click()
'MsgBox PesanQ
'If PesanQ = 1 Then
'    End
'Else
   ' frmMessage.Hide
   ' Unload Me
   ' Load frmxLog ' vbModal
'End If
End Sub


Private Sub LOK_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Shape5.FillColor = vbBlue
End Sub
