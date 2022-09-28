VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmKlasifikasi 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TARIF-NJOPTKP-KELAS TANAH-KELAS BANGUNAN"
   ClientHeight    =   8085
   ClientLeft      =   -15
   ClientTop       =   270
   ClientWidth     =   12195
   ControlBox      =   0   'False
   Icon            =   "frmKlasifikasi.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8085
   ScaleWidth      =   12195
   Begin VB.CheckBox cMassal 
      Caption         =   "Klik Untuk Membuat Nilai Baru Secara Massal"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   75
      TabIndex        =   8
      Top             =   1410
      Width           =   7755
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   45
      Picture         =   "frmKlasifikasi.frx":1CCA
      ScaleHeight     =   300
      ScaleWidth      =   12045
      TabIndex        =   17
      Top             =   90
      Width           =   12045
      Begin VB.CheckBox chPajak 
         BackColor       =   &H8000000D&
         Caption         =   "Pemutakhiran Data"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   3
         Left            =   10245
         TabIndex        =   2
         Top             =   60
         Width           =   1665
      End
      Begin VB.CheckBox chPajak 
         BackColor       =   &H8000000D&
         Caption         =   "Perekaman Data"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   1
         Left            =   6960
         TabIndex        =   0
         Top             =   45
         Value           =   1  'Checked
         Width           =   1500
      End
      Begin VB.CheckBox chPajak 
         BackColor       =   &H8000000D&
         Caption         =   "Penghapusan Data"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   2
         Left            =   8475
         TabIndex        =   1
         Top             =   60
         Width           =   1695
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "KLASIFIKASI TANAH,BANGUNAN DAN TARIF"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Left            =   120
         TabIndex        =   18
         Top             =   60
         Width           =   3645
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   6360
      TabIndex        =   13
      Top             =   7530
      Width           =   990
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   5385
      TabIndex        =   11
      Top             =   7530
      Width           =   990
   End
   Begin MSComctlLib.ListView vBangunan 
      Height          =   5610
      Left            =   45
      TabIndex        =   16
      Top             =   1815
      Width           =   12030
      _ExtentX        =   21220
      _ExtentY        =   9895
      SortKey         =   1
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483642
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   1
      MousePointer    =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   11
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "a"
         Text            =   "No"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   1
         Text            =   "Kode"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Umur Efektif"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "Biaya Pengganti"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Kondisi Bangunan"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Harga Lama"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "Harga Baru"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "xx"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "yy"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "ZZ"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "XX"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Frame Frame5 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Left            =   45
      TabIndex        =   14
      Top             =   315
      Width           =   12030
      Begin VB.ComboBox ccAkhir 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3915
         TabIndex        =   4
         Top             =   150
         Width           =   1560
      End
      Begin VB.ComboBox ccAwal 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1470
         TabIndex        =   3
         Top             =   150
         Width           =   1560
      End
      Begin VB.ComboBox cboNOP 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   7095
         TabIndex        =   5
         Top             =   165
         Width           =   4845
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Klasifikasi"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   6030
         TabIndex        =   24
         Top             =   225
         Width           =   870
      End
      Begin VB.Label Label2 
         Caption         =   "s/d"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   3195
         TabIndex        =   19
         Top             =   240
         Width           =   1365
      End
      Begin VB.Label Label12 
         Caption         =   "Tahun Awal"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   15
         Top             =   195
         Width           =   1260
      End
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   4410
      TabIndex        =   10
      Top             =   7530
      Width           =   990
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Left            =   45
      TabIndex        =   20
      Top             =   840
      Width           =   12030
      Begin VB.TextBox tBumi 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   3
         Left            =   9045
         TabIndex        =   12
         Top             =   165
         Width           =   2865
      End
      Begin VB.TextBox tBumi 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   2
         Left            =   5370
         TabIndex        =   7
         Top             =   165
         Width           =   2460
      End
      Begin VB.TextBox tBumi 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   1470
         TabIndex        =   6
         Top             =   150
         Width           =   2475
      End
      Begin VB.Label Label5 
         Caption         =   "Range_Max"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   4095
         TabIndex        =   23
         Top             =   210
         Width           =   1260
      End
      Begin VB.Label Label4 
         Caption         =   "Range_Min"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   22
         Top             =   195
         Width           =   1260
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "NJOPTKP"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   8010
         TabIndex        =   21
         Top             =   225
         Width           =   810
      End
   End
   Begin VB.Frame Frame2 
      Height          =   495
      Left            =   7815
      TabIndex        =   25
      Top             =   1335
      Width           =   4245
      Begin VB.CommandButton cmdCari 
         Height          =   375
         Left            =   3780
         Picture         =   "frmKlasifikasi.frx":6332
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   90
         Width           =   420
      End
      Begin VB.TextBox tBumi 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   1290
         TabIndex        =   9
         Top             =   105
         Width           =   2520
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Tarif Baru"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   240
         TabIndex        =   26
         Top             =   180
         Width           =   885
      End
   End
   Begin VB.Image Image1 
      Height          =   9180
      Left            =   -60
      Picture         =   "frmKlasifikasi.frx":6FFC
      Stretch         =   -1  'True
      Top             =   -60
      Width           =   12960
   End
End
Attribute VB_Name = "frmKlasifikasi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim jumRek, K1, K2, PBBMin
Dim xxTahun

Private Sub cmdBangunan_Click()
frmOP_Bangunan.Show
End Sub

Private Sub cmdBumi_Click()
frmOP_Tanah.Show
End Sub


Private Sub cboNOP_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
    SendKeys "{tab}"
End If
If InStr("0123456789", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then
    KeyAscii = 0
End If
End Sub

Private Sub cboNOP_LostFocus()
On Error Resume Next
For i = 0 To cboNOP.ListCount - 1
        If (UCase(cboNOP.List(i)) Like "*" + UCase(cboNOP.Text) + "*" = True) Then
            cboNOP.Text = cboNOP.List(i)
            cboNOP_Click
            Exit Sub
        End If
          If i = cboNOP.ListCount - 1 Then
            If UCase(cboNOP.List(i)) Like "*" + UCase(cboNOP.Text) + "*" = False Then
                cboNOP.Text = cboNOP.List(0)
                cboNOP_Click
                Exit Sub
            End If
        End If
    Next
End Sub

Private Sub ccAkhir_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
    SendKeys "{tab}"
End If
If InStr("0123456789", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then
    KeyAscii = 0
End If
End Sub

Private Sub ccAwal_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
    SendKeys "{tab}"
End If
If InStr("0123456789", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then
    KeyAscii = 0
End If
End Sub

Private Sub chPajak_Click(Index As Integer)
On Error GoTo Salah
Select Case Index
Case 1
    If chPajak(1).Value = 1 Then
        chPajak(2).Value = 0
        chPajak(3).Value = 0
        cmdSave.Caption = "&Save"
        bersih
    Else
        If chPajak(2).Value = 0 And chPajak(3).Value = 0 Then
            chPajak(1).Value = 1
        End If
    End If
Case 2
    
   If chPajak(2).Value = 1 Then
    chPajak(1).Value = 0
    chPajak(3).Value = 0
    cmdSave.Caption = "&Delete"
    xTanya = MsgBox("Apa anda yakin menghapus Klasifikasi Tarif ?", vbQuestion + vbYesNo, "Penghapusan")
    
    If xTanya = vbYes Then
        cmdSave.Caption = "&Delete"
      
        bersih
    Else
        chPajak(1).Value = 1
        chPajak(2).Value = 0
        cmdSave.Caption = "&Save"
    End If
    Else
        If chPajak(1).Value = 0 And chPajak(3).Value = 0 Then
            chPajak(2).Value = 1
        End If
   End If
Case 3
    If chPajak(3).Value = 1 Then
        chPajak(1).Value = 0
        chPajak(2).Value = 0
        cmdSave.Caption = "&Update"
        xTanya = MsgBox("Apa anda yakin melakukan pemutakhiran Klasifikasi Tarif?", vbQuestion + vbYesNo, "Pemutakhiran")
        If xTanya = vbYes Then
            cmdSave.Caption = "&Update"
      
            bersih
        Else
            chPajak(1).Value = 1
            chPajak(3).Value = 0
            cmdSave.Caption = "&Save"
        End If
    Else
        If chPajak(1).Value = 0 And chPajak(2).Value = 0 Then
            chPajak(3).Value = 1
        End If
    End If
End Select
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description

End Sub


Private Sub chPajak_KeyPress(Index As Integer, KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub cMassal_Click()
On Error GoTo Salah
If cMassal.Value = 1 Then
For i = 1 To vBangunan.ListItems.Count
    If vBangunan.ListItems.Item(i).ListSubItems(8).Text = "-" Then
        vBangunan.ListItems.Item(i).ListSubItems(5).Text = vBangunan.ListItems.Item(i).ListSubItems(2).Text
        vBangunan.ListItems.Item(i).ListSubItems(6).Text = vBangunan.ListItems.Item(i).ListSubItems(3).Text
        vBangunan.ListItems.Item(i).ListSubItems(7).Text = vBangunan.ListItems.Item(i).ListSubItems(4).Text
        vBangunan.ListItems.Item(i).ListSubItems(5).ForeColor = vbRed
        vBangunan.ListItems.Item(i).ListSubItems(6).ForeColor = vbRed
        vBangunan.ListItems.Item(i).ListSubItems(7).ForeColor = vbRed
    End If
Next
Else
    
    For i = 1 To vBangunan.ListItems.Count
        If vBangunan.ListItems.Item(i).ListSubItems(8).Text = "-" Then
            vBangunan.ListItems.Item(i).ListSubItems(5).Text = 0
            vBangunan.ListItems.Item(i).ListSubItems(6).Text = 0
            vBangunan.ListItems.Item(i).ListSubItems(7).Text = 0
            vBangunan.ListItems.Item(i).ListSubItems(5).ForeColor = vbBlack
            vBangunan.ListItems.Item(i).ListSubItems(6).ForeColor = vbBlack
            vBangunan.ListItems.Item(i).ListSubItems(7).ForeColor = vbBlack
        End If
    Next
End If
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description

End Sub

Private Sub cMassal_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub cmdCancel_Click()
On Error Resume Next
cMassal.Value = 1
For i = 1 To vBangunan.ListItems.Count
    vBangunan.ListItems.Item(i).ListSubItems(8).Text = "-"
Next
cMassal.Value = 0
bersih
End Sub

Private Sub cmdCari_Click()
'If cboNOP.Text = cboNOP.List(0) Then
'    vBangunan.SelectedItem.ListSubItems(4).Text = tBumi(0).Text 'Format(tBumi(0).Text, "#,#0.00")
'    vBangunan.SelectedItem.ListSubItems(5).Text = "OK"
'    vBangunan.SelectedItem.ListSubItems(4).ForeColor = vbBlue
'Else
On Error GoTo Salah
If cboNOP.ListIndex = 0 Then
    vBangunan.SelectedItem.ListSubItems(5).Text = tBumi(1).Text 'Format(tBumi(0).Text, "#,#0.00")
    vBangunan.SelectedItem.ListSubItems(6).Text = tBumi(2).Text 'Format(tBumi(0).Text, "#,#0.00")
    vBangunan.SelectedItem.ListSubItems(7).Text = tBumi(0).Text 'Format(tBumi(0).Text, "#,#0.00")
    vBangunan.SelectedItem.ListSubItems(8).Text = tBumi(3).Text 'Format(tBumi(0).Text, "#,#0.00")
    vBangunan.SelectedItem.ListSubItems(9).Text = "OK"
    vBangunan.SelectedItem.ListSubItems(5).ForeColor = vbBlue
    vBangunan.SelectedItem.ListSubItems(6).ForeColor = vbBlue
    vBangunan.SelectedItem.ListSubItems(7).ForeColor = vbBlue
    vBangunan.SelectedItem.ListSubItems(8).ForeColor = vbBlue
Else
    vBangunan.SelectedItem.ListSubItems(5).Text = tBumi(1).Text 'Format(tBumi(0).Text, "#,#0.00")
    vBangunan.SelectedItem.ListSubItems(6).Text = tBumi(2).Text 'Format(tBumi(0).Text, "#,#0.00")
    vBangunan.SelectedItem.ListSubItems(7).Text = tBumi(0).Text 'Format(tBumi(0).Text, "#,#0.00")
    vBangunan.SelectedItem.ListSubItems(8).Text = "OK"
    vBangunan.SelectedItem.ListSubItems(5).ForeColor = vbBlue
    vBangunan.SelectedItem.ListSubItems(6).ForeColor = vbBlue
    vBangunan.SelectedItem.ListSubItems(7).ForeColor = vbBlue

End If
tBumi(0).Text = 0
vBangunan.SetFocus
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description

End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdID_Click()
On Error Resume Next
xID = 1
frmList_Subjek.Show
End Sub

Private Sub cmdNOP1_Click()
frmNOP.Show
End Sub

Private Sub cmdNOP2_Click()
frmNOP.Show
End Sub

Private Sub cmdNOP3_Click()
frmNOP.Show
End Sub


Private Sub cmdSave_Click()
'On Error GoTo salah
On Error GoTo Salah
Dim Pesan, Judul
If cmdSave.Caption = "&Save" Then
    Pesan = "Apa anda yakin akan menyimpan data ? "
    Judul = "Saved..."
ElseIf cmdSave.Caption = "&Update" Then
    Pesan = "Data yang telah diubah akan disimpan (Update). Lanjutkan? "
    Judul = "Updated..."
Else
    Pesan = "Seluruh record yang tampil akan terhapus. Lanjutkan? "
    Judul = "Deleted..."
End If
TANYA = MsgBox(Pesan, vbInformation + vbYesNo, Judul)
If TANYA = vbYes Then
    If cboNOP.ListIndex = 0 Then
        SIMPAN_tarif
    ElseIf cboNOP.ListIndex = 1 Then
        SIMPAN_KLS_TANAH
    Else
        SIMPAN_KLS_BANGUNAN
    End If
Else
    Exit Sub
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
cboNOP.Clear
cboNOP.AddItem "1 TARIF"
cboNOP.AddItem "2 KELAS BUMI"
cboNOP.AddItem "3 KELAS BANGUNAN"
'cboNOP.AddItem "4 NJOPTKP"
ccAwal.Clear: ccAkhir.Clear
YY = Format(Now, "yyyy")

For i = YY To 1900 Step -1
    ccAwal.AddItem i
    ccAkhir.AddItem i
Next
ccAwal.Text = ccAwal.List(0)
ccAkhir.Text = ccAkhir.List(0)
cMassal.Value = 0
End Sub






Private Sub cboNOP_Click()
On Error GoTo Salah
cMassal.Value = 0
If cboNOP.ListIndex = 0 Then
    CALL_TARIF
    vBangunan.SortKey = 2
    vBangunan.Sorted = True
    vBangunan.Sorted = False
    tBumi(3).Enabled = True
    tBumi(3).BackColor = vbWhite
    Aktif
ElseIf cboNOP.ListIndex = 1 Then
    Aktif
    CALL_BUMI
    tBumi(3).Enabled = False
    tBumi(3).BackColor = vbButtonFace
Else 'If cboNOP.ListIndex = 2 Then
    Aktif
    CALL_BANGUNAN
    tBumi(3).Enabled = False
    tBumi(3).BackColor = vbButtonFace
'Else
'    NONAktif
'    CALL_NJOPTKP
End If
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description

End Sub

Sub CALL_TARIF()
'cboNOP.Clear
On Error GoTo Salah
vBangunan.ListItems.Clear
strK1 = "SELECT * FROM TARIF WHERE THN_AWAL='" & ccAwal.Text & "'" ' AND THN_AKHIR='" & ccAkhir.Text & "'"
openDB (strK1)
If rPajak.RecordCount > 0 Then rPajak.MoveFirst
        i = 0
        Do While Not rPajak.EOF
        i = i + 1
        vBangunan.ListItems.Add i, "", Format(i, "#")
        vBangunan.ListItems.Item(i).ListSubItems.Add 1, "", rPajak!NJOPTKP 'Format(I, "#")
                vBangunan.ListItems.Item(i).ListSubItems.Add 2, "", Format(rPajak!NJOP_MIN, "#,#0.00")
                vBangunan.ListItems.Item(i).ListSubItems.Add 3, "", Format(rPajak!NJOP_MAX, "#,#0.00")
                vBangunan.ListItems.Item(i).ListSubItems.Add 4, "", rPajak!NILAI_TARIF
                vBangunan.ListItems.Item(i).ListSubItems.Add 5, "", 0
                vBangunan.ListItems.Item(i).ListSubItems.Add 6, "", 0
                vBangunan.ListItems.Item(i).ListSubItems.Add 7, "", 0
                vBangunan.ListItems.Item(i).ListSubItems.Add 8, "", 0
                vBangunan.ListItems.Item(i).ListSubItems.Add 9, "", "-"
                vBangunan.ListItems.Item(i).ListSubItems.Add 10, "", "-"
                vBangunan.ColumnHeaders(2).Text = "NJOPTKP"
                vBangunan.ColumnHeaders(3).Text = "NJOP MIN LAMA"
                vBangunan.ColumnHeaders(4).Text = "NJOP MAX LAMA"
                vBangunan.ColumnHeaders(5).Text = "TARIF LAMA"
                vBangunan.ColumnHeaders(6).Text = "NJOP MIN BARU"
                vBangunan.ColumnHeaders(7).Text = "NJOP MAX BARU"
                vBangunan.ColumnHeaders(8).Text = "TARIF BARU"
                vBangunan.ColumnHeaders(9).Text = "NJOPTKP BARU"
                vBangunan.ColumnHeaders(10).Text = "STATUS"
                vBangunan.ColumnHeaders(11).Text = "KET"
                vBangunan.ColumnHeaders(2).Width = 1900
                vBangunan.ColumnHeaders(3).Width = 1900
                vBangunan.ColumnHeaders(4).Width = 1900
                vBangunan.ColumnHeaders(5).Width = 1400
                vBangunan.ColumnHeaders(6).Width = 1900
                vBangunan.ColumnHeaders(7).Width = 1900
                vBangunan.ColumnHeaders(8).Width = 1400
                vBangunan.ColumnHeaders(9).Width = 1900
                vBangunan.ColumnHeaders(3).Alignment = lvwColumnLeft
                vBangunan.ColumnHeaders(4).Alignment = lvwColumnRight
                vBangunan.ColumnHeaders(5).Alignment = lvwColumnRight
                vBangunan.ColumnHeaders(6).Alignment = lvwColumnRight
                vBangunan.ColumnHeaders(7).Alignment = lvwColumnRight
                vBangunan.ColumnHeaders(8).Alignment = lvwColumnRight
                vBangunan.ColumnHeaders(9).Alignment = lvwColumnRight
                vBangunan.ColumnHeaders(10).Alignment = lvwColumnLeft
        'End If
        
        rPajak.MoveNext
        Loop
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description
        
End Sub
Sub CALL_NJOPTKP()
'cboNOP.Clear
On Error GoTo Salah
vBangunan.ListItems.Clear
strK1 = "SELECT * FROM TARIF"
openDB (strK1)
If rPajak.RecordCount > 0 Then rPajak.MoveFirst
        i = 0
        Do While Not rPajak.EOF
        i = i + 1
        vBangunan.ListItems.Add i, "", Format(i, "#")
        vBangunan.ListItems.Item(i).ListSubItems.Add 1, "", Format(i, "#")
                vBangunan.ListItems.Item(i).ListSubItems.Add 2, "", Format(rPajak!NJOP_MIN, "#,#0.00")
                vBangunan.ListItems.Item(i).ListSubItems.Add 3, "", Format(rPajak!NJOP_MAX, "#,#0.00")
                vBangunan.ListItems.Item(i).ListSubItems.Add 4, "", Format(rPajak!NJOPTKP, "#,#0.00")
                vBangunan.ListItems.Item(i).ListSubItems.Add 5, "", 0
                vBangunan.ListItems.Item(i).ListSubItems.Add 6, "", 0
                vBangunan.ListItems.Item(i).ListSubItems.Add 7, "", 0
                vBangunan.ListItems.Item(i).ListSubItems.Add 8, "", "-"
                vBangunan.ListItems.Item(i).ListSubItems.Add 9, "", "-"
                
                vBangunan.ColumnHeaders(3).Text = "NJOP MIN LAMA"
                vBangunan.ColumnHeaders(4).Text = "NJOP MAX LAMA"
                vBangunan.ColumnHeaders(5).Text = "NJOPTKP_LAMA"
                vBangunan.ColumnHeaders(6).Text = "NJOP MIN BARU"
                vBangunan.ColumnHeaders(7).Text = "NJOP MAX BARU"
                vBangunan.ColumnHeaders(8).Text = "NJOPTKP_BARU"
                vBangunan.ColumnHeaders(9).Text = "STATUS"
                vBangunan.ColumnHeaders(10).Text = "KET"
                vBangunan.ColumnHeaders(2).Width = 800
                vBangunan.ColumnHeaders(3).Width = 1900
                vBangunan.ColumnHeaders(4).Width = 1900
                vBangunan.ColumnHeaders(5).Width = 1400
                vBangunan.ColumnHeaders(6).Width = 1900
                vBangunan.ColumnHeaders(7).Width = 1900
                vBangunan.ColumnHeaders(8).Width = 1400
                vBangunan.ColumnHeaders(3).Alignment = lvwColumnLeft
                vBangunan.ColumnHeaders(4).Alignment = lvwColumnRight
                vBangunan.ColumnHeaders(5).Alignment = lvwColumnRight
                vBangunan.ColumnHeaders(6).Alignment = lvwColumnRight
                vBangunan.ColumnHeaders(7).Alignment = lvwColumnRight
                vBangunan.ColumnHeaders(8).Alignment = lvwColumnRight
                vBangunan.ColumnHeaders(9).Alignment = lvwColumnLeft
                vBangunan.ColumnHeaders(10).Alignment = lvwColumnLeft
        'End If
        
        rPajak.MoveNext
        Loop
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description
        
End Sub

Sub CALL_BUMI()
'cboNOP.Clear
On Error GoTo Salah
vBangunan.ListItems.Clear
strK1 = "SELECT * FROM KELAS_TANAH WHERE THN_AWAL_KLS_TANAH='" & ccAwal.Text & "' "
openDB (strK1)
If rPajak.RecordCount > 0 Then rPajak.MoveFirst
        i = 0
        Do While Not rPajak.EOF
        i = i + 1
        vBangunan.ListItems.Add i, "", Format(i, "#")
        vBangunan.ListItems.Item(i).ListSubItems.Add 1, "", rPajak!KD_KLS_TANAH
                'vBangunan.ListItems.Item(I).ListSubItems.Add 2, "", rPajak!NILAI_MIN_TANAH * 1000 & " s.d " & rPajak!NILAI_MAX_TANAH * 1000
                vBangunan.ListItems.Item(i).ListSubItems.Add 2, "", Format(rPajak!NILAI_MIN_TANAH, "#,#0.00")
                vBangunan.ListItems.Item(i).ListSubItems.Add 3, "", Format(rPajak!NILAI_MAX_TANAH, "#,#0.00")
                vBangunan.ListItems.Item(i).ListSubItems.Add 4, "", Format(rPajak!NILAI_PER_M2_TANAH, "#,#0.00")
                vBangunan.ListItems.Item(i).ListSubItems.Add 5, "", 0
                vBangunan.ListItems.Item(i).ListSubItems.Add 6, "", 0
                vBangunan.ListItems.Item(i).ListSubItems.Add 7, "", 0
                vBangunan.ListItems.Item(i).ListSubItems.Add 8, "", "-"
                vBangunan.ListItems.Item(i).ListSubItems.Add 9, "", "-"
                vBangunan.ListItems.Item(i).ListSubItems.Add 10, "", "-"
                
                vBangunan.ColumnHeaders(2).Text = "KELAS"
                vBangunan.ColumnHeaders(3).Text = "RANGE_MIN"
                vBangunan.ColumnHeaders(4).Text = "RANGE_MAX"
                vBangunan.ColumnHeaders(5).Text = "NILAI LAMA"
                vBangunan.ColumnHeaders(6).Text = "R_MIN_BARU"
                vBangunan.ColumnHeaders(7).Text = "R_MAX_BARU"
                vBangunan.ColumnHeaders(8).Text = "NILAI BARU"
                vBangunan.ColumnHeaders(9).Text = "STATUS"
                vBangunan.ColumnHeaders(10).Text = "KET"
                vBangunan.ColumnHeaders(2).Width = 800
                vBangunan.ColumnHeaders(3).Width = 1700
                vBangunan.ColumnHeaders(4).Width = 1700
                vBangunan.ColumnHeaders(5).Width = 1700
                vBangunan.ColumnHeaders(6).Width = 1700
                vBangunan.ColumnHeaders(7).Width = 1700
                vBangunan.ColumnHeaders(8).Width = 1700
                vBangunan.ColumnHeaders(9).Width = 800
                vBangunan.ColumnHeaders(10).Width = 800
                vBangunan.ColumnHeaders(3).Alignment = lvwColumnRight
                vBangunan.ColumnHeaders(4).Alignment = lvwColumnRight
                vBangunan.ColumnHeaders(5).Alignment = lvwColumnRight
                vBangunan.ColumnHeaders(6).Alignment = lvwColumnRight
                vBangunan.ColumnHeaders(7).Alignment = lvwColumnRight
                vBangunan.ColumnHeaders(8).Alignment = lvwColumnRight
        rPajak.MoveNext
        Loop
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description
    
End Sub
Sub CALL_BANGUNAN()
'cboNOP.Clear
On Error GoTo Salah
vBangunan.ListItems.Clear
strK1 = "SELECT * FROM KELAS_BANGUNAN WHERE THN_AWAL_KLS_BNG='" & ccAwal.Text & "' "
openDB (strK1)
If rPajak.RecordCount > 0 Then rPajak.MoveFirst
        i = 0
        Do While Not rPajak.EOF
        i = i + 1
        vBangunan.ListItems.Add i, "", Format(i, "#")
        vBangunan.ListItems.Item(i).ListSubItems.Add 1, "", rPajak!KD_KLS_BNG
                vBangunan.ListItems.Item(i).ListSubItems.Add 2, "", Format(rPajak!NILAI_MIN_BNG, "#,#0.00")
                vBangunan.ListItems.Item(i).ListSubItems.Add 3, "", Format(rPajak!NILAI_MAX_BNG, "#,#0.00")
                vBangunan.ListItems.Item(i).ListSubItems.Add 4, "", Format(rPajak!NILAI_PER_M2_BNG, "#,#0.00")
                vBangunan.ListItems.Item(i).ListSubItems.Add 5, "", 0
                vBangunan.ListItems.Item(i).ListSubItems.Add 6, "", 0
                vBangunan.ListItems.Item(i).ListSubItems.Add 7, "", 0
                vBangunan.ListItems.Item(i).ListSubItems.Add 8, "", "-"
                vBangunan.ListItems.Item(i).ListSubItems.Add 9, "", "-"
                vBangunan.ListItems.Item(i).ListSubItems.Add 10, "", "-"
                
                vBangunan.ColumnHeaders(2).Text = "KELAS"
                vBangunan.ColumnHeaders(3).Text = "RANGE_MIN"
                vBangunan.ColumnHeaders(4).Text = "RANGE_MAX"
                vBangunan.ColumnHeaders(5).Text = "NILAI LAMA"
                vBangunan.ColumnHeaders(6).Text = "R_MIN_BARU"
                vBangunan.ColumnHeaders(7).Text = "R_MAX_BARU"
                vBangunan.ColumnHeaders(8).Text = "NILAI BARU"
                vBangunan.ColumnHeaders(9).Text = "STATUS"
                vBangunan.ColumnHeaders(10).Text = "KET"
                vBangunan.ColumnHeaders(2).Width = 800
                vBangunan.ColumnHeaders(3).Width = 1700
                vBangunan.ColumnHeaders(4).Width = 1700
                vBangunan.ColumnHeaders(5).Width = 1700
                vBangunan.ColumnHeaders(6).Width = 1700
                vBangunan.ColumnHeaders(7).Width = 1700
                vBangunan.ColumnHeaders(8).Width = 1700
                vBangunan.ColumnHeaders(9).Width = 800
                vBangunan.ColumnHeaders(10).Width = 800
                vBangunan.ColumnHeaders(3).Alignment = lvwColumnRight
                vBangunan.ColumnHeaders(4).Alignment = lvwColumnRight
                vBangunan.ColumnHeaders(5).Alignment = lvwColumnRight
                vBangunan.ColumnHeaders(6).Alignment = lvwColumnRight
                vBangunan.ColumnHeaders(7).Alignment = lvwColumnRight
                vBangunan.ColumnHeaders(8).Alignment = lvwColumnRight
        rPajak.MoveNext
        Loop
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description
    
End Sub


Private Sub tBumi_KeyPress(Index As Integer, KeyAscii As Integer)
On Error Resume Next
Select Case Index
    Case 0, 1, 2, 3
        If KeyAscii = 13 Then
            SendKeys "{TAB}"
            KeyAscii = 0
        End If
        If InStr("0123456789.-,", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then
            KeyAscii = 0
        End If
End Select
End Sub



Private Sub vBangunan_Click()
'If vBangunan.SelectedItem.ListSubItems(5).Text = 1 Then
'If cboNOP.Text = cboNOP.List(0) Then
'    For I = 1 To vBangunan.ListItems.Count
'        vBangunan.ListItems.Item(I).ListSubItems(6).Text = "-"
'    Next
'        vBangunan.SelectedItem.ListSubItems(6).Text = "Proses"
'        tBumi(0).Text = vBangunan.SelectedItem.ListSubItems(3).Text
'    For I = 1 To vBangunan.ListItems.Count
'            If vBangunan.ListItems.Item(I).ListSubItems(5).Text = "OK" Then
'                vBangunan.ListItems.Item(I).ListSubItems(5).Text = "OK"
'            Else
'                vBangunan.ListItems.Item(I).ListSubItems(5).Text = "-"
'            End If
'
'    Next
'Else
On Error GoTo Salah
If cboNOP.ListIndex = 0 Then
    For i = 1 To vBangunan.ListItems.Count
        vBangunan.ListItems.Item(i).ListSubItems(10).Text = "-"
    Next
        vBangunan.SelectedItem.ListSubItems(10).Text = "Proses"
        tBumi(0).Text = vBangunan.SelectedItem.ListSubItems(4).Text
        tBumi(2).Text = vBangunan.SelectedItem.ListSubItems(3).Text
        tBumi(1).Text = vBangunan.SelectedItem.ListSubItems(2).Text
        tBumi(3).Text = vBangunan.SelectedItem.ListSubItems(1).Text
    For i = 1 To vBangunan.ListItems.Count
            If vBangunan.ListItems.Item(i).ListSubItems(9).Text = "OK" Then
                vBangunan.ListItems.Item(i).ListSubItems(9).Text = "OK"
            Else
                vBangunan.ListItems.Item(i).ListSubItems(9).Text = "-"
            End If
    Next
Else
        For i = 1 To vBangunan.ListItems.Count
        vBangunan.ListItems.Item(i).ListSubItems(9).Text = "-"
    Next
        vBangunan.SelectedItem.ListSubItems(9).Text = "Proses"
        tBumi(0).Text = vBangunan.SelectedItem.ListSubItems(4).Text
        tBumi(2).Text = vBangunan.SelectedItem.ListSubItems(3).Text
        tBumi(1).Text = vBangunan.SelectedItem.ListSubItems(2).Text
    For i = 1 To vBangunan.ListItems.Count
            If vBangunan.ListItems.Item(i).ListSubItems(8).Text = "OK" Then
                vBangunan.ListItems.Item(i).ListSubItems(8).Text = "OK"
            Else
                vBangunan.ListItems.Item(i).ListSubItems(8).Text = "-"
            End If
    Next

End If
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description

End Sub




Private Sub vBangunan_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error Resume Next
vBangunan.SortKey = ColumnHeader.Index - 1
vBangunan.Sorted = True
vBangunan.Sorted = False
vBangunan.SortOrder = lvwAscending
End Sub


Private Sub vBangunan_KeyDown(KeyCode As Integer, Shift As Integer)
vBangunan_Click
End Sub

Private Sub vBangunan_KeyUp(KeyCode As Integer, Shift As Integer)
 vBangunan_Click
End Sub
Sub NONAktif()
On Error Resume Next
tBumi(1).Enabled = False
tBumi(1).BackColor = vbButtonFace
tBumi(2).Enabled = False
tBumi(2).BackColor = vbButtonFace
Label4.ForeColor = &HE0E0E0
Label5.ForeColor = &HE0E0E0
cMassal.Enabled = False
End Sub
Sub Aktif()
On Error Resume Next
tBumi(1).Enabled = True
tBumi(1).BackColor = vbWhite
tBumi(2).Enabled = True
tBumi(2).BackColor = vbWhite
Label4.ForeColor = vbBlack
Label5.ForeColor = vbBlack
cMassal.Enabled = True
End Sub

Sub bersih()
On Error Resume Next
tBumi(0).Text = 0
tBumi(1).Text = 0
tBumi(2).Text = 0
tBumi(3).Text = 0
vBangunan.ListItems.Clear
End Sub

Sub SIMPAN_tarif()
On Error GoTo Salah
Screen.MousePointer = vbHourglass
QSTR = "SELECT * FROM TARIF WHERE THN_AWAL='" & ccAwal.Text & "'" ' AND THN_AKHIR='" & ccAkhir.Text & "'"
openDB (QSTR)
If rPajak.RecordCount > 0 Then rPajak.MoveFirst
If Not rPajak.EOF Then ' JIKA ADA
    If cmdSave.Caption = "&Save" Then
    MsgBox "Tarif Tahun Awal: " & ccAwal.Text & " Sudah dibuat sebelumnya...", vbCritical, "Data Exist"
            Screen.MousePointer = vbDefault
            Exit Sub
    End If
End If
If chPajak(2).Value = 1 Then GoTo Loncat1
For J = 1 To vBangunan.ListItems.Count
    If (vBangunan.ListItems.Item(J).ListSubItems(6).Text = "" Or vBangunan.ListItems.Item(J).ListSubItems(6).Text = "-" Or vBangunan.ListItems.Item(J).ListSubItems(6).Text = 0) And (vBangunan.ListItems.Item(J).ListSubItems(7).Text <> "OK") Then
        MsgBox "Nilai Belum Lengkap, Silahkan dilengkapi terlebih dahulu...", vbCritical, "Tetnong"
            Screen.MousePointer = vbDefault
            Exit Sub
    End If
Next
Loncat1:
If (cmdSave.Caption = "&Update" And chPajak(3).Value = 1) Or (cmdSave.Caption = "&Delete" And chPajak(2).Value = 1) Then

    Do While Not rPajak.EOF
        rPajak.Delete adAffectCurrent
        rPajak.Update
        rPajak.MoveNext
    Loop
    If cmdSave.Caption = "&Delete" And chPajak(2).Value = 1 Then
        
        GoTo Keluar
    End If
End If
For i = 1 To vBangunan.ListItems.Count
        rPajak.AddNew
        rPajak!KD_PROPINSI = "12"
        rPajak!KD_DATI2 = "12"
        rPajak!THN_AWAL = ccAwal.Text
        rPajak!THN_AKHIR = ccAkhir.Text
        rPajak!NJOP_MIN = vBangunan.ListItems.Item(i).ListSubItems(5).Text
        rPajak!NJOP_MAX = vBangunan.ListItems.Item(i).ListSubItems(6).Text
        rPajak!NILAI_TARIF = vBangunan.ListItems.Item(i).ListSubItems(7).Text
        rPajak!NJOPTKP = vBangunan.ListItems.Item(i).ListSubItems(8).Text
    rPajak.Update
    
Next
Keluar:
vBangunan.ListItems.Clear
'chPajak(1).Value = 1
'cmdSave.Caption = "&Save"
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description

Screen.MousePointer = vbDefault
End Sub
Sub SIMPAN_KLS_TANAH()
On Error GoTo Salah
Screen.MousePointer = vbHourglass
QSTR = "SELECT * FROM KELAS_TANAH WHERE THN_AWAL_KLS_TANAH='" & ccAwal.Text & "'" ' AND THN_AKHIR='" & ccAkhir.Text & "'"
openDB (QSTR)
If rPajak.RecordCount > 0 Then rPajak.MoveFirst
If Not rPajak.EOF Then ' JIKA ADA
    If cmdSave.Caption = "&Save" Then
    MsgBox "Klasifikasi Tanah Tahun Awal: " & ccAwal.Text & " Sudah dibuat sebelumnya...", vbCritical, "Data Exist"
            Screen.MousePointer = vbDefault
            Exit Sub
    End If
End If
If chPajak(2).Value = 1 Then GoTo Loncat1
For J = 1 To vBangunan.ListItems.Count
    If (vBangunan.ListItems.Item(J).ListSubItems(6).Text = "" Or vBangunan.ListItems.Item(J).ListSubItems(6).Text = "-" Or vBangunan.ListItems.Item(J).ListSubItems(6).Text = 0) And (vBangunan.ListItems.Item(J).ListSubItems(7).Text <> "OK") Then
        MsgBox "Nilai Belum Lengkap, Silahkan dilengkapi terlebih dahulu...", vbCritical, "Tetnong"
            Screen.MousePointer = vbDefault
            Exit Sub
    End If
Next
Loncat1:
If (cmdSave.Caption = "&Update" And chPajak(3).Value = 1) Or (cmdSave.Caption = "&Delete" And chPajak(2).Value = 1) Then

    Do While Not rPajak.EOF
        rPajak.Delete adAffectCurrent
        rPajak.Update
        rPajak.MoveNext
    Loop
    If cmdSave.Caption = "&Delete" And chPajak(2).Value = 1 Then
        
        GoTo Keluar
    End If
End If
For i = 1 To vBangunan.ListItems.Count
        rPajak.AddNew
        rPajak!KD_KLS_TANAH = vBangunan.ListItems.Item(i).ListSubItems(1).Text
        rPajak!THN_AWAL_KLS_TANAH = ccAwal.Text
        rPajak!THN_AKHIR_KLS_TANAH = ccAkhir.Text
        rPajak!NILAI_MIN_TANAH = vBangunan.ListItems.Item(i).ListSubItems(5).Text
        rPajak!NILAI_MAX_TANAH = vBangunan.ListItems.Item(i).ListSubItems(6).Text
        rPajak!NILAI_PER_M2_TANAH = vBangunan.ListItems.Item(i).ListSubItems(7).Text
    rPajak.Update
    
Next
Keluar:
vBangunan.ListItems.Clear
'chPajak(1).Value = 1
'cmdSave.Caption = "&Save"
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description

Screen.MousePointer = vbDefault
End Sub


Sub SIMPAN_KLS_BANGUNAN()
On Error GoTo Salah
Screen.MousePointer = vbHourglass
QSTR = "SELECT * FROM KELAS_BANGUNAN WHERE THN_AWAL_KLS_BNG='" & ccAwal.Text & "'" ' AND THN_AKHIR='" & ccAkhir.Text & "'"
openDB (QSTR)
If rPajak.RecordCount > 0 Then rPajak.MoveFirst
If Not rPajak.EOF Then ' JIKA ADA
    If cmdSave.Caption = "&Save" Then
    MsgBox "Klasifikasi Tanah Tahun Awal: " & ccAwal.Text & " Sudah dibuat sebelumnya...", vbCritical, "Data Exist"
            Screen.MousePointer = vbDefault
            Exit Sub
    End If
End If
If chPajak(2).Value = 1 Then GoTo Loncat1
For J = 1 To vBangunan.ListItems.Count
    If (vBangunan.ListItems.Item(J).ListSubItems(6).Text = "" Or vBangunan.ListItems.Item(J).ListSubItems(6).Text = "-" Or vBangunan.ListItems.Item(J).ListSubItems(6).Text = 0) And (vBangunan.ListItems.Item(J).ListSubItems(7).Text <> "OK") Then
        MsgBox "Nilai Belum Lengkap, Silahkan dilengkapi terlebih dahulu...", vbCritical, "Tetnong"
            Screen.MousePointer = vbDefault
            Exit Sub
    End If
Next
Loncat1:
If (cmdSave.Caption = "&Update" And chPajak(3).Value = 1) Or (cmdSave.Caption = "&Delete" And chPajak(2).Value = 1) Then

    Do While Not rPajak.EOF
        rPajak.Delete adAffectCurrent
        rPajak.Update
        rPajak.MoveNext
    Loop
    If cmdSave.Caption = "&Delete" And chPajak(2).Value = 1 Then
        
        GoTo Keluar
    End If
End If
For i = 1 To vBangunan.ListItems.Count
        rPajak.AddNew
        
        rPajak!KD_KLS_BNG = vBangunan.ListItems.Item(i).ListSubItems(1).Text
        rPajak!THN_AWAL_KLS_BNG = ccAwal.Text
        rPajak!THN_AKHIR_KLS_BNG = ccAkhir.Text
        rPajak!NILAI_MIN_BNG = vBangunan.ListItems.Item(i).ListSubItems(5).Text
        rPajak!NILAI_MAX_BNG = vBangunan.ListItems.Item(i).ListSubItems(6).Text
        rPajak!NILAI_PER_M2_BNG = vBangunan.ListItems.Item(i).ListSubItems(7).Text
    rPajak.Update
    
Next
Keluar:
vBangunan.ListItems.Clear
'chPajak(1).Value = 1
'cmdSave.Caption = "&Save"
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description

Screen.MousePointer = vbDefault
End Sub



