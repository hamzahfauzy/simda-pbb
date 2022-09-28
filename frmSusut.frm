VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSusut 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Range Penyusutan Nilai Bangunan"
   ClientHeight    =   8085
   ClientLeft      =   -15
   ClientTop       =   270
   ClientWidth     =   12195
   ControlBox      =   0   'False
   Icon            =   "frmSusut.frx":0000
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
      Height          =   270
      Left            =   90
      TabIndex        =   15
      Top             =   900
      Width           =   12000
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   60
      Picture         =   "frmSusut.frx":1CCA
      ScaleHeight     =   300
      ScaleWidth      =   11985
      TabIndex        =   7
      Top             =   90
      Width           =   11985
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
         Left            =   10290
         TabIndex        =   18
         Top             =   30
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
         Left            =   7005
         TabIndex        =   17
         Top             =   30
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
         Left            =   8520
         TabIndex        =   16
         Top             =   30
         Width           =   1695
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Umur Efektif dan Penyusutan"
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
         Left            =   390
         TabIndex        =   8
         Top             =   60
         Width           =   2415
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
      TabIndex        =   0
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
      TabIndex        =   1
      Top             =   7530
      Width           =   990
   End
   Begin MSComctlLib.ListView vBangunan 
      Height          =   6255
      Left            =   60
      TabIndex        =   6
      Top             =   1155
      Width           =   12030
      _ExtentX        =   21220
      _ExtentY        =   11033
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
         Name            =   "Arial Unicode MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   9
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
         Object.Width           =   4057
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "Harga Baru"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   7
         Text            =   "Status"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Ket"
         Object.Width           =   1764
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
      Height          =   585
      Left            =   45
      TabIndex        =   3
      Top             =   315
      Width           =   5235
      Begin VB.ComboBox cboNOP 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   2
         Left            =   1125
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   180
         Width           =   4035
      End
      Begin VB.Label Label12 
         Caption         =   "Kode Range"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   60
         TabIndex        =   5
         Top             =   210
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
      TabIndex        =   2
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
      Height          =   585
      Left            =   5280
      TabIndex        =   9
      Top             =   315
      Width           =   6855
      Begin VB.CommandButton cmdCari 
         Height          =   345
         Left            =   6405
         Picture         =   "frmSusut.frx":6332
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   165
         Width           =   375
      End
      Begin VB.TextBox tBumi 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   4155
         TabIndex        =   12
         Top             =   165
         Width           =   2250
      End
      Begin VB.TextBox tBumi 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   1050
         TabIndex        =   10
         Top             =   165
         Width           =   2130
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Harga Baru"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3240
         TabIndex        =   14
         Top             =   225
         Width           =   810
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Range_Min"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   105
         TabIndex        =   11
         Top             =   225
         Width           =   795
      End
   End
   Begin VB.Image Image1 
      Height          =   9180
      Left            =   -120
      Picture         =   "frmSusut.frx":6FFC
      Stretch         =   -1  'True
      Top             =   -60
      Width           =   12945
   End
End
Attribute VB_Name = "frmSusut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim jumRek, K1, K2, PBBMin
Private Sub cmdBangunan_Click()
frmOP_Bangunan.Show
End Sub

Private Sub cmdBumi_Click()
frmOP_Tanah.Show
End Sub


Private Sub chPajak_Click(Index As Integer)
On Error GoTo Salah
Select Case Index
Case 1
    If chPajak(1).Value = 1 Then
        chPajak(2).Value = 0
        chPajak(3).Value = 0
        cmdSave.Caption = "&Save"
        xxTahun = (cboNOP(1).Text * 1) - 1
        'Bersih
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
    xTanya = MsgBox("Apa anda yakin menghapus Nilai DBKB Material?", vbQuestion + vbYesNo, "Penghapusan")
    
    If xTanya = vbYes Then
        cmdSave.Caption = "&Delete"
        xxTahun = cboNOP(1).Text * 1
        'Bersih
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
        xTanya = MsgBox("Apa anda yakin melakukan pemutakhiran Nilai DBKB Material?", vbQuestion + vbYesNo, "Pemutakhiran")
        If xTanya = vbYes Then
            cmdSave.Caption = "&Update"
            xxTahun = cboNOP(1).Text * 1
'            Bersih
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


Private Sub cMassal_Click()
On Error GoTo Salah
If cMassal.Value = 1 Then
    For i = 1 To vBangunan.ListItems.Count
    If vBangunan.ListItems.Item(i).ListSubItems(7).Text = "-" Then
        If cboNOP(2).Text = cboNOP(2).List(0) Then
            vBangunan.ListItems.Item(i).ListSubItems(5).Text = vBangunan.ListItems.Item(i).ListSubItems(3).Text
            vBangunan.ListItems.Item(i).ListSubItems(6).Text = vBangunan.ListItems.Item(i).ListSubItems(4).Text
            vBangunan.ListItems.Item(i).ListSubItems(5).ForeColor = vbRed
        Else
            vBangunan.ListItems.Item(i).ListSubItems(6).Text = vBangunan.ListItems.Item(i).ListSubItems(5).Text
        End If
            vBangunan.ListItems.Item(i).ListSubItems(6).ForeColor = vbRed
    End If
    Next
Else
    For i = 1 To vBangunan.ListItems.Count
    If vBangunan.ListItems.Item(i).ListSubItems(7).Text = "-" Then
        If cboNOP(2).Text = cboNOP(2).List(0) Then
            vBangunan.ListItems.Item(i).ListSubItems(5).Text = 0
            vBangunan.ListItems.Item(i).ListSubItems(6).Text = 0
            vBangunan.ListItems.Item(i).ListSubItems(5).ForeColor = vbBlack
        Else
            vBangunan.ListItems.Item(i).ListSubItems(6).Text = 0
        End If
            vBangunan.ListItems.Item(i).ListSubItems(6).ForeColor = vbBlack
    End If
    Next
End If
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description
End Sub

Private Sub cmdCancel_Click()
On Error GoTo Salah
cMassal.Value = 0
For i = 1 To vBangunan.ListItems.Count
    vBangunan.ListItems.Item(i).ListSubItems(7).Text = "-"
Next
cMassal.Value = 0
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description
End Sub

Private Sub cmdCari_Click()
On Error Resume Next
If cboNOP(2).Text = cboNOP(2).List(0) Then
    vBangunan.SelectedItem.ListSubItems(5).Text = Format(tBumi(1).Text, "#,#0.00")
    vBangunan.SelectedItem.ListSubItems(6).Text = Format(tBumi(0).Text, "#,#0.00")
Else
    vBangunan.SelectedItem.ListSubItems(6).Text = Format(tBumi(0).Text, "#,#0.00")
End If
    vBangunan.SelectedItem.ListSubItems(7).Text = "OK"

tBumi(0).Text = 0
vBangunan.SetFocus
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

Private Sub Form_Activate()
On Error Resume Next
Me.Top = (frmUtama.ScaleHeight - Me.Height) / 2
    Me.Left = (frmUtama.Width - Me.Width) / 2
End Sub



Private Sub cboNOP_Click(Index As Integer)
'On Error Resume Next
On Error GoTo Salah
Screen.MousePointer = vbHourglass
Dim xKon
Select Case Index
    Case 0
            cboNOP(2).Clear
    Case 1
         cboNOP(2).Clear
    Case 2
        cMassal.Value = 0
        JFAS = Left(cboNOP(2).Text, 2)
        'If JFAS = "02" Or JFAS = "07" Or JFAS = "09" Then JFAS = "02"
        vBangunan.ListItems.Clear
        'If cboNOP(2).ListIndex = cboNOP(2).ListCount - 1 Then
        If cboNOP(2).Text = cboNOP(2).List(0) Then
        tBumi(1).Enabled = True
        tBumi(1).BackColor = vbWhite
        Label4.ForeColor = vbBlack '&H00E0E0E0&
        Label1.Caption = "Range_Max"
        STRITEM = "SELECT * FROM RANGE_PENYUSUTAN ORDER BY KD_RANGE_PENYUSUTAN ASC"
        openDB (STRITEM)

        If rPajak.RecordCount > 0 Then rPajak.MoveFirst
        i = 0
        Do While Not rPajak.EOF
        i = i + 1
        vBangunan.ListItems.Add i, "", Format(i, "#")
        vBangunan.ListItems.Item(i).ListSubItems.Add 1, "", Format(i, "#")
                vBangunan.ListItems.Item(i).ListSubItems.Add 2, "", rPajak!KD_RANGE_PENYUSUTAN
                vBangunan.ListItems.Item(i).ListSubItems.Add 3, "", Format(Trim(rPajak!NILAI_MIN_PENYUSUTAN), "#,#0.00")
                vBangunan.ListItems.Item(i).ListSubItems.Add 4, "", Format(Trim(rPajak!NILAI_MAX_PENYUSUTAN), "#,#0.00")
                vBangunan.ListItems.Item(i).ListSubItems.Add 5, "", 0
                vBangunan.ListItems.Item(i).ListSubItems.Add 6, "", 0
                vBangunan.ListItems.Item(i).ListSubItems.Add 7, "", "-"
                vBangunan.ListItems.Item(i).ListSubItems.Add 8, "", "-"
                
                vBangunan.ColumnHeaders(3).Text = "KODE"
                vBangunan.ColumnHeaders(4).Text = "NILAI_MIN_LAMA"
                vBangunan.ColumnHeaders(5).Text = "NILAI_MAX_LAMA"
                vBangunan.ColumnHeaders(6).Text = "NILAI_MIN_BARU"
                vBangunan.ColumnHeaders(7).Text = "NILAI_MAX_BARU"
                vBangunan.ColumnHeaders(2).Width = 0
                vBangunan.ColumnHeaders(3).Width = 800
                vBangunan.ColumnHeaders(4).Width = 2200
                vBangunan.ColumnHeaders(5).Width = 1900
                vBangunan.ColumnHeaders(6).Width = 1900
                vBangunan.ColumnHeaders(7).Width = 1900
                vBangunan.ColumnHeaders(3).Alignment = lvwColumnRight
                vBangunan.ColumnHeaders(4).Alignment = lvwColumnRight
                vBangunan.ColumnHeaders(5).Alignment = lvwColumnRight
                vBangunan.ColumnHeaders(6).Alignment = lvwColumnRight
                vBangunan.ColumnHeaders(7).Alignment = lvwColumnRight
        'End If
        
        rPajak.MoveNext
        Loop
        tBumi(0).Text = 0
        '===================
        Else
        tBumi(1).Enabled = False
        tBumi(1).BackColor = vbButtonFace
        Label4.ForeColor = &HE0E0E0
        Label1.Caption = "Harga Baru"
        STRITEM = "SELECT PENYUSUTAN.UMUR_EFEKTIF, PENYUSUTAN.KD_RANGE_PENYUSUTAN, PENYUSUTAN.KONDISI_BNG_SUSUT, RANGE_PENYUSUTAN.NILAI_MIN_PENYUSUTAN, RANGE_PENYUSUTAN.NILAI_MAX_PENYUSUTAN, PENYUSUTAN.NILAI_PENYUSUTAN FROM PENYUSUTAN INNER JOIN RANGE_PENYUSUTAN ON PENYUSUTAN.KD_RANGE_PENYUSUTAN = RANGE_PENYUSUTAN.KD_RANGE_PENYUSUTAN where PENYUSUTAN.KD_RANGE_PENYUSUTAN='" & Trim(Left(cboNOP(2).Text, 2)) & "' order by PENYUSUTAN.UMUR_EFEKTIF ASC"
        openDB (STRITEM)

        If rPajak.RecordCount > 0 Then rPajak.MoveFirst
        i = 0
        Do While Not rPajak.EOF
        i = i + 1
        vBangunan.ListItems.Add i, "", Format(i, "#")
        vBangunan.ListItems.Item(i).ListSubItems.Add 1, "", Format(i, "#")
                vBangunan.ListItems.Item(i).ListSubItems.Add 2, "", Format(rPajak!UMUR_EFEKTIF, "00")
                vBangunan.ListItems.Item(i).ListSubItems.Add 3, "", Trim(rPajak![NILAI_MIN_PENYUSUTAN]) & " s.d " & rPajak!NILAI_MAX_PENYUSUTAN
                If rPajak!KONDISI_BNG_SUSUT = "1" Then
                    xKon = "Sangat Baik"
                ElseIf rPajak!KONDISI_BNG_SUSUT = "2" Then
                    xKon = "Baik"
                ElseIf rPajak!KONDISI_BNG_SUSUT = "3" Then
                    xKon = "Sedang"
                Else
                    xKon = "Jelek"
                End If
                vBangunan.ListItems.Item(i).ListSubItems.Add 4, "", rPajak!KONDISI_BNG_SUSUT & " - " & xKon
                vBangunan.ListItems.Item(i).ListSubItems.Add 5, "", Format(Trim(rPajak!NILAI_PENYUSUTAN), "#,#0.00")
                vBangunan.ListItems.Item(i).ListSubItems.Add 6, "", 0
                vBangunan.ListItems.Item(i).ListSubItems.Add 7, "", "-"
                vBangunan.ListItems.Item(i).ListSubItems.Add 8, "", "-"
                
                vBangunan.ColumnHeaders(3).Text = "UMUR EFEKTIF"
                vBangunan.ColumnHeaders(4).Text = "RANGE SUSUT"
                vBangunan.ColumnHeaders(5).Text = "KONDISI BANGUNAN"
                vBangunan.ColumnHeaders(6).Text = "NILAI LAMA"
                vBangunan.ColumnHeaders(7).Text = "NILAI BARU"
                vBangunan.ColumnHeaders(2).Width = 0
                vBangunan.ColumnHeaders(3).Width = 1000
                vBangunan.ColumnHeaders(4).Width = 2200
                vBangunan.ColumnHeaders(5).Width = 1600
                
                vBangunan.ColumnHeaders(3).Alignment = lvwColumnRight
                vBangunan.ColumnHeaders(4).Alignment = lvwColumnCenter
                vBangunan.ColumnHeaders(5).Alignment = lvwColumnLeft
                vBangunan.ColumnHeaders(6).Alignment = lvwColumnRight
                vBangunan.ColumnHeaders(7).Alignment = lvwColumnRight
        'End If
        rPajak.MoveNext
        Loop
        tBumi(0).Text = 0
        End If
    Case 3
        
        'tBumi(0).Text = K1 & "." & K2 & "." & Left(cboNOP(0).Text, 3) & "." & Left(cboNOP(1).Text, 3) & "." & Left(cbonop(1).Text, 3) & "-" & cbonop(2).Text & "." & Left(cboNOP(4).Text, 1)
End Select
vBangunan.SortKey = 2
vBangunan.Sorted = True
vBangunan.Sorted = False
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description

Screen.MousePointer = vbDefault
End Sub

Private Sub cboNOP_DropDown(Index As Integer)
On Error Resume Next
Select Case Index
    Case 1
    Case 2
    
    callKomp1
    Case 3
        
        
End Select

End Sub
Sub callKomp1()
On Error GoTo Salah
cboNOP(2).Clear ': cboZNT.Clear
'strK1 = "Select * From GROUP_RESOURCE order by KD_GROUP_RESOURCE ASC"
'strK1 = "SELECT PENYUSUTAN.UMUR_EFEKTIF, PENYUSUTAN.KD_RANGE_PENYUSUTAN, PENYUSUTAN.KONDISI_BNG_SUSUT, RANGE_PENYUSUTAN.NILAI_MIN_PENYUSUTAN, RANGE_PENYUSUTAN.NILAI_MAX_PENYUSUTAN, PENYUSUTAN.NILAI_PENYUSUTAN FROM PENYUSUTAN INNER JOIN RANGE_PENYUSUTAN ON PENYUSUTAN.KD_RANGE_PENYUSUTAN = RANGE_PENYUSUTAN.KD_RANGE_PENYUSUTAN"
cboNOP(2).AddItem 1 & " RANGE PENYUSUTAN"
strK1 = "SELECT * FROM RANGE_PENYUSUTAN ORDER BY KD_RANGE_PENYUSUTAN"
openDB (strK1)
If rPajak.RecordCount > 0 Then rPajak.MoveFirst
i = 1
Do While Not rPajak.EOF
    i = i + 1
        cboNOP(2).AddItem i & " BIAYA PENGGANTI BARU " & rPajak!KD_RANGE_PENYUSUTAN 'rPajak!NILAI_MIN_PENYUSUTAN & " " & rPajak!NILAI_MAX_PENYUSUTAN
rPajak.MoveNext
Loop

If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description
End Sub


Private Sub vBangunan_Click()
'If vBangunan.SelectedItem.ListSubItems(5).Text = 1 Then
On Error GoTo Salah
If cboNOP(2).Text = cboNOP(2).List(0) Then
    tBumi(1).Text = vBangunan.SelectedItem.ListSubItems(3).Text
    tBumi(0).Text = vBangunan.SelectedItem.ListSubItems(4).Text
Else
    tBumi(0).Text = vBangunan.SelectedItem.ListSubItems(5).Text
End If
    For i = 1 To vBangunan.ListItems.Count
        vBangunan.ListItems.Item(i).ListSubItems(8).Text = "-"
    Next
 vBangunan.SelectedItem.ListSubItems(8).Text = "Proses"

'    tBumi(1).Text = vBangunan.SelectedItem.ListSubItems(5).Text
    For i = 1 To vBangunan.ListItems.Count
            If vBangunan.ListItems.Item(i).ListSubItems(7).Text = "OK" Then
                vBangunan.ListItems.Item(i).ListSubItems(7).Text = "OK"
            Else
                vBangunan.ListItems.Item(i).ListSubItems(7).Text = "-"
            End If
            
    Next
 
            If vBangunan.SelectedItem.ListSubItems(7).Text = "OK" Then
                Exit Sub
                vBangunan.SetFocus
            Else
               ' tBumi(0).SetFocus
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
