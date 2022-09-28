VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmxLog 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Run APP!"
   ClientHeight    =   2475
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8805
   ControlBox      =   0   'False
   Icon            =   "frmxLog.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2475
   ScaleWidth      =   8805
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame6 
      BackColor       =   &H8000000D&
      Height          =   2625
      Left            =   0
      TabIndex        =   8
      Top             =   -90
      Width           =   2190
      Begin VB.Frame Frame4 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         Height          =   480
         Left            =   165
         TabIndex        =   15
         Top             =   1440
         Width           =   1830
         Begin VB.Label LOFF 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "LOCAL !"
            BeginProperty Font 
               Name            =   "Arial Rounded MT Bold"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   300
            Left            =   60
            MousePointer    =   99  'Custom
            TabIndex        =   16
            Top             =   135
            Width           =   1680
         End
         Begin VB.Shape Shape4 
            BorderColor     =   &H00000000&
            FillColor       =   &H00C0C0C0&
            FillStyle       =   0  'Solid
            Height          =   420
            Left            =   30
            Shape           =   4  'Rounded Rectangle
            Top             =   30
            Width           =   1770
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         Height          =   480
         Left            =   165
         TabIndex        =   13
         Top             =   2025
         Width           =   1830
         Begin VB.Label LBatal 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "KELUAR !"
            BeginProperty Font 
               Name            =   "Arial Rounded MT Bold"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   300
            Left            =   90
            MousePointer    =   99  'Custom
            TabIndex        =   14
            Top             =   135
            Width           =   1695
         End
         Begin VB.Shape Shape3 
            BorderColor     =   &H00000000&
            FillColor       =   &H00C0C0C0&
            FillStyle       =   0  'Solid
            Height          =   420
            Left            =   30
            Shape           =   4  'Rounded Rectangle
            Top             =   30
            Width           =   1770
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         Height          =   480
         Left            =   150
         TabIndex        =   11
         Top             =   855
         Width           =   1845
         Begin VB.Label LCLIENT 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "CLIENT !"
            BeginProperty Font 
               Name            =   "Arial Rounded MT Bold"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   300
            Left            =   105
            MousePointer    =   99  'Custom
            TabIndex        =   12
            Top             =   120
            Width           =   1680
         End
         Begin VB.Shape Shape1 
            BorderColor     =   &H00000000&
            FillColor       =   &H00C0C0C0&
            FillStyle       =   0  'Solid
            Height          =   420
            Left            =   30
            Shape           =   4  'Rounded Rectangle
            Top             =   30
            Width           =   1785
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   480
         Left            =   150
         TabIndex        =   9
         Top             =   285
         Width           =   1845
         Begin VB.Label LCSERVER 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "SERVER !"
            BeginProperty Font 
               Name            =   "Arial Rounded MT Bold"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   300
            Left            =   90
            MousePointer    =   99  'Custom
            TabIndex        =   10
            Top             =   120
            Width           =   1695
         End
         Begin VB.Shape Shape2 
            BorderColor     =   &H00000000&
            FillColor       =   &H00C0C0C0&
            FillStyle       =   0  'Solid
            Height          =   420
            Left            =   30
            Shape           =   4  'Rounded Rectangle
            Top             =   30
            Width           =   1785
         End
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00000000&
      Caption         =   "Connect"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   2505
      Left            =   2265
      TabIndex        =   1
      Top             =   0
      Width           =   6465
      Begin VB.CommandButton cOK 
         Caption         =   "OK"
         Height          =   495
         Left            =   4800
         TabIndex        =   21
         Top             =   2850
         Width           =   945
      End
      Begin VB.TextBox tNama 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   2040
         TabIndex        =   17
         Top             =   945
         Width           =   3795
      End
      Begin VB.DriveListBox Drive1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2040
         TabIndex        =   5
         Top             =   315
         Width           =   4305
      End
      Begin VB.TextBox tFolder 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   2040
         TabIndex        =   6
         Top             =   630
         Width           =   4305
      End
      Begin VB.TextBox tLOK 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   75
         Locked          =   -1  'True
         MousePointer    =   99  'Custom
         MultiLine       =   -1  'True
         TabIndex        =   3
         Top             =   1320
         Width           =   6330
      End
      Begin VB.Image Image5 
         Appearance      =   0  'Flat
         Height          =   255
         Left            =   6210
         MousePointer    =   99  'Custom
         Picture         =   "frmxLog.frx":2EFA
         Stretch         =   -1  'True
         ToolTipText     =   "Keluar dari Form"
         Top             =   30
         Width           =   255
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   ".mdb"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   5835
         MousePointer    =   99  'Custom
         TabIndex        =   20
         Top             =   945
         Width           =   510
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama File Database"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   180
         Left            =   240
         MousePointer    =   99  'Custom
         TabIndex        =   19
         Top             =   1005
         Width           =   1665
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Folder"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   180
         Left            =   240
         MousePointer    =   99  'Custom
         TabIndex        =   18
         Top             =   675
         Width           =   1065
      End
      Begin VB.Label LOK 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Create NOW!"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   300
         Left            =   2700
         MousePointer    =   99  'Custom
         TabIndex        =   7
         Top             =   1995
         Width           =   1275
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Lokasi Database"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   180
         Left            =   240
         MousePointer    =   99  'Custom
         TabIndex        =   4
         Top             =   360
         Width           =   1380
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Lokasi DATABASE !"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   270
         Left            =   2460
         MousePointer    =   99  'Custom
         TabIndex        =   2
         Top             =   855
         Width           =   1770
      End
      Begin VB.Shape Shape5 
         BorderColor     =   &H00000000&
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00C0C0C0&
         FillStyle       =   0  'Solid
         Height          =   360
         Left            =   2640
         Shape           =   4  'Rounded Rectangle
         Top             =   1950
         Width           =   1395
      End
      Begin VB.Shape Shape6 
         BorderColor     =   &H00FFFFFF&
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   450
         Left            =   2595
         Shape           =   4  'Rounded Rectangle
         Top             =   1905
         Width           =   1470
      End
   End
   Begin VB.DirListBox Dir1 
      Height          =   1665
      Left            =   4725
      TabIndex        =   0
      Top             =   3090
      Visible         =   0   'False
      Width           =   3810
   End
   Begin MSComDlg.CommonDialog gLOK 
      Left            =   4245
      Top             =   2235
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmxLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim dbCall As New ADODB.Connection
Dim dbSIMPEG As New ADODB.Connection
Dim rsAdm As New ADODB.Recordset
Dim pLOG, xxFolder, xLen, xxLok, xxAdmin, xL1, xL2, xL3, xL4
Sub accLok()
On Error Resume Next
    Open App.Path & "\LokFile.txt" For Input As #1
    Line Input #1, xL1
    Line Input #1, xL2
    Line Input #1, xL3
    Line Input #1, xL4
    Close #1
End Sub

Private Sub cOK_Click()
LOK_Click
End Sub

Private Sub cOK_GotFocus()
Shape5.FillColor = vbBlue
End Sub

'Private Sub cFol_Click()
'nFol = InputBox("Nama Folder :", "Create Folder")
'MkDir Dir2.Path & "\" & nFol
'End Sub

'Private Sub Dir2_Change()
'File1.FileName = Dir2.Path
'tFolder.Text = Dir2.Path
'End Sub

Private Sub Drive1_Change()
On Error GoTo xSalah
'If Dir.Visible = True Then
'Dir2.Path = Drive1.Drive
'Else
Dir1.Path = Drive1.Drive
'End If
'tFolder.Text = left(Drive1.Drive, 2) & "\dsnCServer"
tLOK.Text = ""
tLOK.Text = left(Drive1.Drive, 2) & "\" & Trim(tFolder.Text) & "\dbSIMPEG\" & Trim(tNama.Text) & ".mdb"
xSalah:
'On Error Resume Next
If Err.Number = 68 Or Err.Number = 86 Then
    MsgBox "Ganti drive lain...", vbCritical, "REJECT"
    X = Drive1.List(Drive1.ListIndex - 1)
    Drive1.Drive = X
    'tFolder.Text = left(Drive1.Drive, 2) & "\dsnCServer"
    Exit Sub
End If
End Sub


Private Sub Drive1_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
    tFolder.SetFocus
End If
End Sub

'Private Sub File1_Click()
'tFolder.Text = Dir2.Path & "\" & File1.FileName
'End Sub

Private Sub Form_Activate()
On Error Resume Next
Me.tOp = (Screen.Height - Me.Height) / 2
Me.left = (Screen.Width - Me.Width) / 2
End Sub

Private Sub Form_Load()
On Error Resume Next
Dim cLogin, cFolder

Me.Height = 2955
Me.Width = 2280
accLok
LokasiAdmin = Trim(xL1)
LokasiData = Trim(xL2)
LokasiFoto = Trim(xL3)
jenisAPP = Val(Trim(xL4))
If LokasiAdmin = "" Or LokasiData = "" Or LokasiFoto = "" Or jenisAPP = "" Then
    Exit Sub
Else
    If jenisAPP = 1 Then
        t1 = GetAttr(App.Path & "\dsnCServer\simpegPB.dsn")
    ElseIf jenisAPP = 2 Then
        t1 = GetAttr(App.Path & "\dsnCLient\simpegPB.dsn")
    Else
        GoTo kl
    End If
    
    If (t1 <> 32 Or Err.Number = 53) Then
        Exit Sub
    End If
kl:
    xxadm = GetAttr(LokasiAdmin)
    xxdb = GetAttr(LokasiData)
    If (xxadm <> 32 Or xxdb <> 32 Or Err.Number = 53) And jenisAPP <> 0 Then
    
        PesanQ = 1
        'Unload frmxLog
        'MsgBox 1
        'frmMessage.Show vbModal
        'PesanQ = 0
    'End If
    
    
    'If (xxDB <> 32 Or Err.Number = 53) And jenisAPP <> 0 Then
        
     '   PesanQ = 2
       ' Me.Hide
       'MsgBox PesanQ
        frmMessage.Show vbModal
       ' PesanQ = 0
        Frame5.left = 0
        Frame6.Visible = False
        Kill App.Path & "\dsnCLient\" & "*.dsn"
        LCLIENT_Click
        Me.Height = 2955
        Me.Width = Frame5.Width + 100
        
        Exit Sub
    End If
    
    If cUser = 1 Then Exit Sub
    
    frmLogin.Show
    Unload Me
End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Shape1.FillColor = &HC0C0C0
Shape2.FillColor = &HC0C0C0
Shape3.FillColor = &HC0C0C0
Shape4.FillColor = &HC0C0C0
Shape5.FillColor = &HC0C0C0

End Sub


Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
If PesanQ = 1 Then 'Unload Me 'End
'frmMessage.Show
Unload frmMessage
Unload Me
cUser = 0
End If
End Sub

Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Shape1.FillColor = &HC0C0C0
Shape2.FillColor = &HC0C0C0
Shape3.FillColor = &HC0C0C0
Shape4.FillColor = &HC0C0C0
Shape5.FillColor = &HC0C0C0

End Sub
Private Sub Frame5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Shape1.FillColor = &HC0C0C0
Shape2.FillColor = &HC0C0C0
Shape3.FillColor = &HC0C0C0
Shape4.FillColor = &HC0C0C0
Shape5.FillColor = &HC0C0C0
tLOK.FontBold = False
tLOK.ForeColor = vbWhite
End Sub
Private Sub Frame2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Shape1.FillColor = &HC0C0C0
Shape2.FillColor = &HC0C0C0
Shape3.FillColor = &HC0C0C0
Shape4.FillColor = &HC0C0C0
Shape5.FillColor = &HC0C0C0

End Sub

Private Sub Frame3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Shape1.FillColor = &HC0C0C0
Shape2.FillColor = &HC0C0C0
Shape3.FillColor = &HC0C0C0
Shape4.FillColor = &HC0C0C0
Shape5.FillColor = &HC0C0C0

End Sub
Private Sub Frame4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Shape1.FillColor = &HC0C0C0
Shape2.FillColor = &HC0C0C0
Shape3.FillColor = &HC0C0C0
Shape4.FillColor = &HC0C0C0
Shape5.FillColor = &HC0C0C0

End Sub

Private Sub Image5_Click()
Unload Me
End Sub

Private Sub LBatal_Click()
On Error Resume Next
If LBatal.Caption = "KELUAR !" Then
    If cUser = 0 Then
        End
    Else
        Unload Me
        Exit Sub
    End If
Else
Me.Height = 2955
Me.Width = 2280
    
    LBatal.Caption = "KELUAR !"
End If
pLOG = 0
Me.tOp = (Screen.Height - Me.Height) / 2
Me.left = (Screen.Width - Me.Width) / 2

End Sub

Private Sub LBatal_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Shape3.FillColor = vbRed
End Sub

Private Sub LCLIENT_Click()
On Error Resume Next
LBatal.Caption = "B A T A L !"
Me.Height = 2955
Me.Width = 8835
Drive1.Visible = False
Label2.Visible = False
Label1.Visible = True
tNama.Visible = False
Label3.Visible = False
Label4.Visible = False
Label5.Visible = False
tLOK.Width = 6285
tLOK.Alignment = 2
tLOK.left = 150
tFolder.Visible = False
tLOK.Locked = True
tLOK.Text = "[KLIK disini untuk menghubungkan CLIENT ke SERVER!]"
Frame5.Caption = "[Connect to SERVER]"
Me.tOp = (Screen.Height - Me.Height) / 2
Me.left = (Screen.Width - Me.Width) / 2
pLOG = 2
LOK.Caption = "Connect"
End Sub

Private Sub LCLIENT_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Shape1.FillColor = vbGreen
End Sub

Private Sub LCSERVER_Click()
On Error Resume Next
LBatal.Caption = "B A T A L !"
'Label1.Caption = "Nama Folder!"
Me.Height = 2955
Me.Width = 8835
pLOG = 1
tLOK.Text = "" 'App.Path & "\dbsimpeg.mdb"
Drive1.Visible = True
Label2.Visible = True
Label3.Visible = True
Label4.Visible = True
Label5.Visible = True
Label1.Visible = False
tFolder.Visible = True
tNama.Visible = True
tLOK.Width = 4410
tLOK.Alignment = 0
tLOK.left = 2010
'tLOK.Locked = False
'tLOK.Text = left(Drive1.Drive, 2) & "\dsnCServer"
Frame5.Caption = "[Membuat Nama Folder Database SERVER !]"
Drive1.SetFocus
Me.tOp = (Screen.Height - Me.Height) / 2
Me.left = (Screen.Width - Me.Width) / 2
tFolder.Text = "dsnCServer"
tNama.Text = "dbSIMPEG"
tLOK.Text = left(Drive1.Drive, 2) & "\" & Trim(tFolder.Text) & "\dbSIMPEG\" & Trim(tNama.Text) & ".mdb"
LOK.Caption = "Create NOW!"
'Me.Caption = Screen.Width ' Me.left
'Me.left = Me.left - Me.Width
End Sub

Private Sub LCSERVER_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Shape2.FillColor = vbGreen
End Sub


Private Sub LIns_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Shape6.FillColor = vbRed
End Sub

Private Sub LOFF_Click()
On Error Resume Next
HapusUsr
pLOG = 0
tLOK.Text = App.Path & "\dbSIMPEG\dbSIMPEG.mdb"
LokasiDB
accLok
LokasiAdmin = Trim(xL1)
LokasiData = Trim(xL2)
LokasiFoto = Trim(xL3)
jenisAPP = Val(Trim(xL4))

Unload Me
frmLogin.Show
End Sub

Private Sub LOFF_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Shape4.FillColor = vbGreen
End Sub

Private Sub LOK_Click()
On Error Resume Next
If tLOK = "" Or tLOK.Text = "[KLIK disini untuk menghubungkan CLIENT ke SERVER!]" Then MsgBox "Lokasi Database SERVER Belum ada...!", vbCritical, "Disconnect": Exit Sub

'---------------------
HapusUsr
LokasiDB
If pLOG = 1 Then
    conSR
ElseIf pLOG = 2 Then
    conCL
End If

accLok
LokasiAdmin = Trim(xL1)
LokasiData = Trim(xL2)
LokasiFoto = Trim(xL3)
jenisAPP = Val(Trim(xL4))
If PesanQ = 1 Then
    frmLogin.Show
Else
    Unload Me
    frmLogin.Show
End If
PesanQ = 0
End Sub

Private Sub LOK_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Shape5.FillColor = vbBlue
End Sub



Private Sub tFolder_Change()
On Error Resume Next
tLOK.Text = left(Drive1.Drive, 2) & "\" & Trim(tFolder.Text) & "\dbSIMPEG\" & Trim(tNama.Text) & ".mdb"
End Sub


Private Sub tFolder_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 13 Then
    tNama.SetFocus
    tFolder.Text = Trim(tFolder.Text)
End If
End Sub

 

Private Sub tFolder_KeyPress(KeyAscii As Integer)
On Error Resume Next
If InStr("/\:*?|<>'", Chr(KeyAscii)) <> 0 Then 'Or KeyAscii <> vbKeyBack Then
   KeyAscii = 0
End If
End Sub

Private Sub tFolder_LostFocus()
On Error Resume Next
If tFolder.Text = "" Then
    tLOK.Text = left(Drive1.Drive, 2) & "\dsnCServer\dbSIMPEG\" & Trim(tNama.Text) & ".mdb"
End If
End Sub

Private Sub tLOK_Click()
On Error Resume Next
If pLOG = 1 Then Exit Sub
 With gLOK
        .DialogTitle = "Pilih Lokasi Database...!"
        .CancelError = False
        'ToDo: set the flags and attributes of the common dialog control
        '.Filter = "All Files (*.*)|*.*"
        .Filter = "File Access(*.MDB)|*.MDB"
        .ShowOpen
        If Len(.FileName) = 0 Then
            Exit Sub
        End If
        xLen = Len(.FileTitle)
        xxLok = left(.FileName, Len(.FileName) - xLen - 1)
        xxAdmin = xxLok & "\Admin.mdb"
    tLOK.Text = .FileName
    
    End With
    cOK.SetFocus
End Sub

Sub conCL()
On Error Resume Next
xDSN = GetAttr(App.Path & "\dsnCLient")
If xDSN <> 16 Or Err.Number = 53 Then
    MkDir App.Path & "\dsnCLient"
End If

xDSN1 = GetAttr(App.Path & "\dsnCLient\adminPB.dsn")
If xDSN1 <> 32 Then 'Err.Number = 53 Then
    Dim i As Integer
    Open App.Path & "\adminPB.txt" For Output As #1
    Print #1, "[ODBC]"
    Print #1, "DRIVER=Microsoft Access Driver (*.mdb)"
    Print #1, "UID=admin"
    Print #1, "UserCommitSync=Yes"
    Print #1, "Threads=3"
    Print #1, "SafeTransactions=0"
    Print #1, "PageTimeout=5"
    Print #1, "MaxScanRows=8"
    Print #1, "MaxBufferSize=2048"
    Print #1, "FIL=MS Access"
    Print #1, "DriverId=25"
    Print #1, "DefaultDir=" & xxLok
    Print #1, "DBQ=" & xxAdmin
    Close #1
    FileCopy App.Path & "\adminPB.txt", App.Path & "\dsnCLient\adminPB.dsn"

End If
'    xd1 = GetAttr(xxAdmin)
'    If xd1 <> 32 Then 'Or Err.Number = 53 Then
'        FileCopy App.Path & "\Admin.mdb", xxAdmin
'    End If

xxDSN = GetAttr(App.Path & "\dsnCLient\simpegPB.dsn")

If xxDSN <> 32 Then 'Err.Number = 53 Then
    Open App.Path & "\simpegPB.txt" For Output As #1
    Print #1, "[ODBC]"
    Print #1, "DRIVER=Microsoft Access Driver (*.mdb)"
    Print #1, "UID=admin"
    Print #1, "UserCommitSync=Yes"
    Print #1, "Threads=3"
    Print #1, "SafeTransactions=0"
    Print #1, "PageTimeout=5"
    Print #1, "MaxScanRows=8"
    Print #1, "MaxBufferSize=2048"
    Print #1, "FIL=MS Access"
    Print #1, "DriverId=25"
    Print #1, "DefaultDir=" & xxLok
    Print #1, "DBQ=" & tLOK.Text
    Close #1
    FileCopy App.Path & "\simpegPB.txt", App.Path & "\dsnCLient\simpegPB.dsn"
    'FileCopy App.Path & "\dbSIMPEG.mdb", "C:\dsnCLient\dbSIMPEG\dbSIMPEG.mdb"
End If
'    xd2 = GetAttr(tLOK.Text)
'    If xd2 <> 32 Then 'Or Err.Number = 53 Then
'        FileCopy App.Path & "\dbSIMPEG.mdb", tLOK.Text
'    End If


End Sub
Sub conSR()
On Error Resume Next
Dim nmFolder
nmFolder = left(Drive1.Drive, 2) & "\" & Trim(tFolder.Text)
'xDSN = GetAttr(tLOK.Text)
'If xDSN <> 16 Or Err.Number = 53 Then
'    MkDir tLOK.Text
'    MkDir tLOK.Text & "\dbSIMPEG"
'    MkDir tLOK.Text & "\dbSIMPEG\Foto"
'End If
'----Repair for up command
xDSN = GetAttr(nmFolder)
If xDSN <> 16 Or Err.Number = 53 Then
    MkDir nmFolder
    MkDir nmFolder & "\dbSIMPEG"
    MkDir nmFolder & "\dbSIMPEG\Foto"
End If
xDSN1 = GetAttr(App.Path & "\dsnCServer")
If xDSN1 <> 16 Or Err.Number = 53 Then
    MkDir App.Path & "\dsnCServer"
End If

xDSN2 = GetAttr(App.Path & "\dsnCServer\adminPB.dsn")
If xDSN2 <> 32 Then 'Err.Number = 53 Then
    Dim i As Integer
    Open App.Path & "\adminPB.txt" For Output As #1
    Print #1, "[ODBC]"
    Print #1, "DRIVER=Microsoft Access Driver (*.mdb)"
    Print #1, "UID=admin"
    Print #1, "UserCommitSync=Yes"
    Print #1, "Threads=3"
    Print #1, "SafeTransactions=0"
    Print #1, "PageTimeout=5"
    Print #1, "MaxScanRows=8"
    Print #1, "MaxBufferSize=2048"
    Print #1, "FIL=MS Access"
    Print #1, "DriverId=25"
    Print #1, "DefaultDir=" & nmFolder & "\dbSIMPEG"
    Print #1, "DBQ=" & nmFolder & "\dbSIMPEG\Admin.mdb"
    Close #1
    FileCopy App.Path & "\adminPB.txt", App.Path & "\dsnCServer\adminPB.dsn"
    xxDSN2 = GetAttr(nmFolder & "\dbSIMPEG\Admin.mdb")
    If xxDSN2 <> 32 Then
        FileCopy App.Path & "\Admin.mdb", nmFolder & "\dbSIMPEG\Admin.mdb"
    End If


End If

xDSN3 = GetAttr(App.Path & "\dsnCServer\simpegPB.dsn")
If xDSN3 <> 32 Then 'Err.Number = 53 Then
    Open App.Path & "\simpegPB.txt" For Output As #1
    Print #1, "[ODBC]"
    Print #1, "DRIVER=Microsoft Access Driver (*.mdb)"
    Print #1, "UID=admin"
    Print #1, "UserCommitSync=Yes"
    Print #1, "Threads=3"
    Print #1, "SafeTransactions=0"
    Print #1, "PageTimeout=5"
    Print #1, "MaxScanRows=8"
    Print #1, "MaxBufferSize=2048"
    Print #1, "FIL=MS Access"
    Print #1, "DriverId=25"
    Print #1, "DefaultDir=" & nmFolder & "\dbSIMPEG"
    Print #1, "DBQ=" & tLOK.Text '& "\dbSIMPEG\dbSIMPEG.mdb"
    Close #1
    FileCopy App.Path & "\simpegPB.txt", App.Path & "\dsnCServer\simpegPB.dsn"
    xxDSN3 = GetAttr(tLOK.Text)
    If xxDSN3 <> 32 Then
        FileCopy App.Path & "\dbSIMPEG.mdb", tLOK.Text ' & "\dbSIMPEG\dbSIMPEG.mdb"
    End If
End If


End Sub
Sub LokasiDB()
On Error Resume Next
Dim nmFolder
nmFolder = left(Drive1.Drive, 2) & "\" & Trim(tFolder.Text)

    Open App.Path & "\LokFile.txt" For Output As #1
    If pLOG = 1 Then
        Print #1, nmFolder & "\dbSIMPEG\Admin.mdb"
        Print #1, tLOK.Text '& "\dbSIMPEG\dbSIMPEG.mdb"
        Print #1, nmFolder & "\dbSIMPEG\Foto\"
    ElseIf pLOG = 2 Then
        Print #1, xxAdmin
        Print #1, tLOK.Text
        Print #1, xxLok & "\Foto\"
    Else
        Print #1, App.Path & "\Admin.mdb"
        Print #1, tLOK.Text
        Print #1, App.Path & "\Foto\"
    End If
        Print #1, pLOG
    Close #1
End Sub

Private Sub tLOK_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    If pLOG = 1 Then Exit Sub
    tLOK_Click
    cOK.SetFocus
End If
End Sub

Private Sub tLOK_KeyPress(KeyAscii As Integer)
On Error Resume Next
If InStr("/\", Chr(KeyAscii)) <> 0 Then 'Or KeyAscii <> vbKeyBack Then
   KeyAscii = 0
End If
End Sub

Private Sub tLOK_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
tLOK.FontBold = True
tLOK.ForeColor = vbBlue
End Sub

Private Sub tNama_Change()
On Error Resume Next
tLOK.Text = left(Drive1.Drive, 2) & "\" & Trim(tFolder.Text) & "\dbSIMPEG\" & Trim(tNama.Text) & ".mdb"
End Sub

Private Sub tNama_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    cOK.SetFocus
End If
End Sub

Private Sub tNama_KeyPress(KeyAscii As Integer)
On Error Resume Next
If InStr("/\:*?|<>'", Chr(KeyAscii)) <> 0 Then 'Or KeyAscii <> vbKeyBack Then
   KeyAscii = 0
End If
End Sub

Private Sub tNama_LostFocus()
On Error Resume Next
If tNama.Text = "" Then
    tLOK.Text = left(Drive1.Drive, 2) & "\dsnCServer\dbSIMPEG\dbSIMPEG.mdb"
End If

End Sub
Sub HapusUsr()
On Error Resume Next
If cUser = 1 Then
Kill App.Path & "\dsnClient\*.*"
RmDir App.Path & "\dsnClient"
Kill App.Path & "\dsnCServer\*.*"
RmDir App.Path & "\dsnCServer"
FileCopy App.Path & "\xLokFile.txt", App.Path & "\LokFile.txt"
End If
cUser = 0
End Sub
