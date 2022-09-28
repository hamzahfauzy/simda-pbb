VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmLogin 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "!Login User...."
   ClientHeight    =   2805
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5415
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2805
   ScaleWidth      =   5415
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox chServer 
      Appearance      =   0  'Flat
      BackColor       =   &H80000003&
      Caption         =   "Cek Server"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   225
      Left            =   165
      TabIndex        =   15
      Top             =   2475
      Width           =   1380
   End
   Begin VB.ComboBox lstComputers 
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1785
      TabIndex        =   2
      Top             =   1245
      Width           =   3375
   End
   Begin MSWinsockLib.Winsock wSock 
      Left            =   1080
      Top             =   1785
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000003&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   15
      ScaleHeight     =   360
      ScaleWidth      =   5385
      TabIndex        =   9
      Top             =   15
      Width           =   5385
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Login..."
         BeginProperty Font 
            Name            =   "Tekton Pro Ext"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   270
         Left            =   15
         TabIndex        =   13
         Top             =   45
         Width           =   810
      End
   End
   Begin VB.Frame Frame4 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1620
      Left            =   2565
      TabIndex        =   3
      Top             =   3495
      Width           =   5535
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "User Name"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   885
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Password"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   4
         Top             =   585
         Width           =   765
      End
   End
   Begin VB.TextBox tUser 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000014&
      BeginProperty Font 
         Name            =   "Calibri Light"
         Size            =   9.75
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   330
      IMEMode         =   3  'DISABLE
      Left            =   1785
      TabIndex        =   0
      Top             =   540
      Width           =   3360
   End
   Begin VB.TextBox tPas 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000014&
      BeginProperty Font 
         Name            =   "Marlett"
         Size            =   9.75
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   1785
      PasswordChar    =   "h"
      TabIndex        =   1
      Top             =   900
      Width           =   3360
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000003&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   450
      Left            =   15
      ScaleHeight     =   450
      ScaleWidth      =   5385
      TabIndex        =   10
      Top             =   2340
      Width           =   5385
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Lokasi Server"
      BeginProperty Font 
         Name            =   "Tekton Pro"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   270
      Left            =   285
      TabIndex        =   14
      Top             =   1245
      Width           =   1140
   End
   Begin VB.Image iLogin 
      Height          =   405
      Left            =   2790
      MouseIcon       =   "frmLogin.frx":1CCA
      MousePointer    =   99  'Custom
      Picture         =   "frmLogin.frx":1FD4
      Stretch         =   -1  'True
      ToolTipText     =   "Login"
      Top             =   1635
      Width           =   360
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   3300
      MouseIcon       =   "frmLogin.frx":11706
      MousePointer    =   99  'Custom
      Picture         =   "frmLogin.frx":11A10
      ToolTipText     =   "Keluar"
      Top             =   1620
      Width           =   480
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H8000000D&
      Height          =   2340
      Left            =   0
      Top             =   0
      Width           =   5415
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "Tekton Pro"
         Size            =   12.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   315
      Left            =   300
      TabIndex        =   12
      Top             =   885
      Width           =   1065
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Username"
      BeginProperty Font 
         Name            =   "Tekton Pro"
         Size            =   12.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   315
      Left            =   315
      TabIndex        =   11
      Top             =   510
      Width           =   1035
   End
   Begin VB.Label LPas 
      Caption         =   "Label1"
      Height          =   300
      Left            =   0
      TabIndex        =   8
      Top             =   2865
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label LLevel 
      Caption         =   "Label3"
      Height          =   285
      Left            =   30
      TabIndex        =   7
      Top             =   3300
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Label LSKPD 
      Caption         =   "Label1"
      Height          =   345
      Left            =   360
      TabIndex        =   6
      Top             =   2430
      Visible         =   0   'False
      Width           =   1965
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H8000000D&
      Height          =   2430
      Left            =   0
      Top             =   375
      Width           =   5415
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ccIP, ccHost, ccPort
Dim intIDX As Integer
Dim ServerList As ListOfServer

Private Sub cmdBatal_Click()
On Error Resume Next
Unload Me
End Sub

Private Sub chServer_Click()
On Error GoTo Salah
If chServer.Value = 1 Then
    Label5.Visible = True
    lstComputers.Visible = True
    lstComputers.Clear
    ServerList = EnumServer(SRV_TYPE_ALL)
    If ServerList.Init Then
        For intIDX = 1 To UBound(ServerList.List)
            'If UCase(ServerList.List(intIDX).ServerName) <> UCase(ccHost) Then ' Komputer current tidak tampil
                lstComputers.AddItem ServerList.List(intIDX).ServerName
            'End If
        Next
    End If
    lstComputers.Text = lstComputers.List(0)
        iLogin.Top = 1740
        Image1.Top = 1710 '1530
        'chServer.Left = 315
Else
    Label5.Visible = False
    lstComputers.Visible = False
            iLogin.Top = 1545
        Image1.Top = 1530
        'chServer.Left = 1785
End If
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description

End Sub

Private Sub Form_Activate()
On Error Resume Next

Me.Top = (frmUtama.ScaleHeight - Me.Height) / 2
Me.Left = (frmUtama.Width - Me.Width) / 2
'Panggil Jenis Penggunaan Kayu Ulin
nonAktifMenu
'MsgBox "IP Address = " & wSock.LocalIP & _
        vbCrLf & "Hostname = " & wSock.LocalHostName & _
        vbCrLf & "Port Number = " & wSock.LocalPort
 'ccIP = wSock.LocalIP
 ccHost = wSock.LocalHostName
 'ccPort = wSock.LocalPort
 lstComputers.Clear
 ServerList = EnumServer(SRV_TYPE_ALL)
    
    ' Loop through all the computers and add them to the listbox
    If ServerList.Init Then
        For intIDX = 1 To UBound(ServerList.List)
            'If UCase(ServerList.List(intIDX).ServerName) <> UCase(ccHost) Then ' Komputer current tidak tampil
                lstComputers.AddItem ServerList.List(intIDX).ServerName
            'End If
        Next
    End If
    lstComputers.Text = lstComputers.List(0)
'chServer.Left = 1785
    If lstComputers.Visible = True Then
        iLogin.Top = 1740
        Image1.Top = 1710 '1530
    Else
        iLogin.Top = 1545
        Image1.Top = 1530
    End If
    If lstComputers.ListCount = 0 Then
        chServer.Visible = False
    Else
        chServer.Visible = True
       ' chServer.Left = 315
    End If
If lstComputers.Text = "" Then lstComputers.Text = ccHost
End Sub

Private Sub iExit_Click()
On Error Resume Next
Unload Me
End Sub

Private Sub Form_Load()
On Error Resume Next
accLok
If nCom = "" Then
    Label5.Visible = True
    lstComputers.Visible = True
    iLogin.Top = 1545
    Image1.Top = 1530
    chServer.Visible = False
Else
    Label5.Visible = False
    lstComputers.Visible = False
    iLogin.Top = 1740
    Image1.Top = 1710 '1530
    chServer.Visible = True
End If
End Sub

'Private Sub Form_Resize()
''On Error Resume Next
'Me.Width = 6420
'Me.Height = 1965
'Me.Top = (mnUtama.ScaleHeight - Me.Height) / 2
'Me.Left = (mnUtama.Width - Me.Width) / 2
'End Sub

Private Sub iLogin_Click()
On Error GoTo Salah
Screen.MousePointer = vbHourglass
accLok
If nCom = "" Or lstComputers.Visible = True Then
    ciptaFile
End If
accLok
            C_ADM = "SELECT * FROM [USERS] WHERE [USERNAME]='" & tUser.Text & "' AND [PASSWORD]='" & tPas.Text & "'"
            openDB (C_ADM)
            
            If Not rPajak.EOF Then
                LPas.Caption = rPajak!Password
                LLevel.Caption = rPajak!WEWENANG
                frmUtama.stsBar.Panels.Item(6).Text = rPajak![UserName] & ": " & rPajak![WEWENANG]
'                    If rsAdm!Level = 1 Then
'                        mnUtama.mnManage.Enabled = True
'                        'mnUtama.mnServer.Enabled = True
'                         mnUtama.icoUser.Enabled = True
'                         mnUtama.mnDefault.Enabled = True
'                         'mnUtama.mnSetting.Enabled = True
'                        'mnUtama.mnClient.Enabled = True
'                    Else
'                        mnUtama.mnDefault.Enabled = False
'                       ' mnUtama.mnSetting.Enabled = False
'                        mnUtama.mnManage.Enabled = False
'                       ' mnUtama.mnServer.Enabled = False
'                        mnUtama.icoUser.Enabled = False
'                        'mnUtama.mnClient.Enabled = False
'
'                    End If
                'rSKPD = rsAdm!SKPD
                If UCase(LPas.Caption) <> UCase(tPas.Text) Then
                    MsgBox "Username atau Password SALAH!", vbCritical, "Error"
                    'Call PESAN(1, "Faild", "username atau password salah", "Silahkan diulangi...!")
                    tPas.SetFocus
                    tPas.Text = ""
                    nonAktifMenu
                    Screen.MousePointer = vbDefault
                    Exit Sub
                End If
            Else
                    MsgBox "Username atau Password SALAH!", vbCritical, "Error"
                     'Call PESAN(1, "Faild", "username atau password salah", "Silahkan diulangi...!")
                    tUser.Text = ""
                    tPas.Text = ""
                    tUser.SetFocus
                    nonAktifMenu
                    Screen.MousePointer = vbDefault
                    Exit Sub
            End If
      

Jadi = 0
uLevel = LLevel.Caption
SenFileRpt

AktifMenu
xLoad = 0
frmUtama.stsBar.Panels.Item(7).Text = UCase(nCom) '" SERVER : " & UCase(nCom)
CALL_ULIN
Unload Me
Salah:
If Err.Number = 0 Or Err.Number = 53 Then GoTo Keluar
MsgBox Err.Number & ": Login GAGAL!", vbCritical, "Fail"
    Open App.Path & "\LokSERVER.txt" For Output As #1
        Print #1, ""
    Close #1
Label5.Visible = True
lstComputers.Visible = True
iLogin.Top = 1740
        Image1.Top = 1710

'mnUtama.Show
Keluar:
Screen.MousePointer = vbDefault
End Sub

Private Sub Image1_Click()
On Error Resume Next
If xLoad = 0 Then
    
    AktifMenu

End If
    Unload Me
                Kill "D:\*.TMP"
                Kill "C:\*.TMP"
                Kill App.Path & "\*.TMP"
                Kill App.Path & "\DATABASE\*.TMP"
                Kill App.Path & "\*.TMP"
                Kill App.Path & "\foto\*.tmp"
'                End
            
End Sub

Private Sub lstComputers_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
    SendKeys "{tab}"
    iLogin_Click
    KeyAscii = 0
End If

End Sub

Private Sub tPas_GotFocus()
On Error Resume Next
tPas.SelStart = 0
  tPas.SelLength = Len(tPas.Text)
End Sub

Private Sub tPas_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
    iLogin_Click
    KeyAscii = 0
End If

End Sub

Private Sub tUser_GotFocus()
On Error Resume Next
tUser.SelStart = 0
  tUser.SelLength = Len(tUser.Text)
End Sub



Private Sub tUser_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
    SendKeys "{tab}"
    KeyAscii = 0
End If
If InStr("'", Chr(KeyAscii)) = 1 Then
  KeyAscii = 0
End If
End Sub

Sub SenFileRpt()
On Error Resume Next
Screen.MousePointer = vbHourglass

GetAttr ("C:\Windows\System32\mscomct2.ocx")
'If Err.Number <> 53 Then GoTo SFR1
FileCopy App.Path & "\mscomct2.ocx", "C:\WINDOWS\system32\mscomct2.ocx"
'SFR1:
GetAttr ("C:\WINDOWS\system32\Crystl32.OCX")
'If Err.Number <> 53 Then GoTo SFR2
FileCopy App.Path & "\Crystl32.OCX", "C:\WINDOWS\system32\Crystl32.OCX"
'SFR2:
GetAttr ("C:\WINDOWS\system32\MSCOMCTL.OCX")
'If Err.Number <> 53 Then GoTo sfr3
FileCopy App.Path & "\MSCOMCTL.OCX", "C:\WINDOWS\system32\MSCOMCTL.OCX"
'sfr3:
GetAttr ("C:\WINDOWS\system32\MSCOMCTL.OCX")
'If Err.Number <> 53 Then Exit Sub
FileCopy App.Path & "\p2sodbc.dll", "C:\WINDOWS\system32\p2sodbc.dll"
Screen.MousePointer = vbHourglass
End Sub



Sub AktifMenu()
On Error Resume Next
'mnUtama.mnMaster.Enabled = True
'mnUtama.mnCari.Enabled = True
'mnUtama.mnReport.Enabled = True
''mnUtama.mnSetting.Enabled = True
'mnUtama.mnHelp.Enabled = True
'mnUtama.Logo2.Enabled = True
'mnUtama.picLogo.Enabled = True
'mnUtama.Logo2.Visible = True
'mnUtama.picLogo.Visible = True
''mnUtama.mnClient.Enabled = True
frmUtama.mnFile.Enabled = True
frmUtama.mnDaftar.Enabled = True
frmUtama.mnNilai.Enabled = True
frmUtama.mnPenetapan.Enabled = True
frmUtama.mnReferensi.Enabled = True
frmUtama.mnLaporan.Enabled = True
frmUtama.mnUtility.Enabled = True
'frmUtama.mnHelp.Enabled = True
frmUtama.mnOFF1.Visible = False
End Sub
Sub nonAktifMenu()
On Error Resume Next
'mnUtama.mnMaster.Enabled = False
'mnUtama.mnCari.Enabled = False
'mnUtama.mnReport.Enabled = False
''mnUtama.mnSetting.Enabled = False
'mnUtama.mnHelp.Enabled = False
'mnUtama.Logo2.Enabled = False
'mnUtama.picLogo.Enabled = False
'mnUtama.Logo2.Visible = False
'mnUtama.picLogo.Visible = False
'mnUtama.mnDefault.Enabled = False
'mnUtama.mnManage.Enabled = False
''mnUtama.mnServer.Enabled = False
''mnUtama.mnClient.Enabled = False
frmUtama.mnFile.Enabled = False
frmUtama.mnDaftar.Enabled = False
frmUtama.mnNilai.Enabled = False
frmUtama.mnPenetapan.Enabled = False
frmUtama.mnReferensi.Enabled = False
frmUtama.mnLaporan.Enabled = False
frmUtama.mnUtility.Enabled = False
'frmUtama.mnHelp.Enabled = False
frmUtama.mnOFF1.Visible = True
End Sub

Sub accLok()
On Error Resume Next
    Open App.Path & "\LokSERVER.txt" For Input As #1
    Line Input #1, nCom
    'Line Input #1, ccHost
    'Line Input #1, ccPort
    Close #1
End Sub
Sub ciptaFile()
On Error Resume Next
   Open App.Path & "\LokSERVER.txt" For Output As #1
        Print #1, UCase(lstComputers.Text)
        'Print #1, ccHost
        'Print #1, ccPort
    Close #1
End Sub
Sub CALL_ULIN()
On Error GoTo Salah
C_STR = "Select * From KAYU_ULIN ORDER BY THN_STATUS_KAYU_ULIN DESC"
openDB (C_STR)
If rPajak.RecordCount > 0 Then rPajak.MoveFirst
ck_Ulin = rPajak!STATUS_KAYU_ULIN
tck_ulin = rPajak!THN_STATUS_KAYU_ULIN
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description

End Sub
