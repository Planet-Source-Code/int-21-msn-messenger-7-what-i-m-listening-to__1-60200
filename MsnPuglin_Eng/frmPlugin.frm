VERSION 5.00
Begin VB.Form frmPlugin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "What I'm Listening To - Plugin"
   ClientHeight    =   4335
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   8895
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPlugin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4335
   ScaleWidth      =   8895
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame FrActivar 
      Height          =   495
      Left            =   3000
      TabIndex        =   16
      Top             =   3600
      Width           =   1575
      Begin VB.Image ImgActivar 
         Height          =   240
         Left            =   120
         Picture         =   "frmPlugin.frx":08CA
         Top             =   160
         Width           =   240
      End
      Begin VB.Label lbActivar 
         AutoSize        =   -1  'True
         Caption         =   "Activate"
         Height          =   195
         Left            =   480
         MouseIcon       =   "frmPlugin.frx":0C54
         MousePointer    =   99  'Custom
         TabIndex        =   17
         Tag             =   "0"
         Top             =   165
         Width           =   690
      End
   End
   Begin VB.Frame frCerrar 
      Height          =   495
      Left            =   6360
      TabIndex        =   20
      Top             =   3600
      Width           =   1575
      Begin VB.Label lbCerrar 
         AutoSize        =   -1  'True
         Caption         =   "Close"
         Height          =   195
         Left            =   480
         MouseIcon       =   "frmPlugin.frx":0DA6
         MousePointer    =   99  'Custom
         TabIndex        =   21
         Tag             =   "0"
         Top             =   165
         Width           =   480
      End
      Begin VB.Image ImgCerrar 
         Height          =   225
         Left            =   120
         Picture         =   "frmPlugin.frx":0EF8
         Top             =   195
         Width           =   210
      End
   End
   Begin VB.Frame frAplicar 
      Height          =   495
      Left            =   4680
      TabIndex        =   18
      Top             =   3600
      Width           =   1575
      Begin VB.Image ImgAplicar 
         Height          =   120
         Left            =   120
         Picture         =   "frmPlugin.frx":11CE
         Top             =   240
         Width           =   150
      End
      Begin VB.Label lbAplicar 
         AutoSize        =   -1  'True
         Caption         =   "Apply"
         Enabled         =   0   'False
         Height          =   195
         Left            =   480
         MouseIcon       =   "frmPlugin.frx":1219
         MousePointer    =   99  'Custom
         TabIndex        =   19
         Tag             =   "0"
         Top             =   165
         Width           =   480
      End
   End
   Begin VB.CheckBox chkWMedia 
      Caption         =   "Window Media"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6240
      TabIndex        =   15
      Top             =   360
      Width           =   1575
   End
   Begin VB.CheckBox chkWinamp 
      Caption         =   "Winamp"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3000
      TabIndex        =   14
      Top             =   360
      Value           =   1  'Checked
      Width           =   1095
   End
   Begin VB.Timer tmCheck 
      Enabled         =   0   'False
      Interval        =   30000
      Left            =   10200
      Top             =   4560
   End
   Begin VB.Frame FrPersonalizado 
      Caption         =   "Custom Format"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   3000
      TabIndex        =   1
      Top             =   840
      Width           =   5895
      Begin VB.TextBox txtFormato 
         Height          =   285
         Left            =   120
         TabIndex        =   9
         Text            =   "{1} by {0}"
         Top             =   2040
         Width           =   5655
      End
      Begin VB.Label lbFormato 
         AutoSize        =   -1  'True
         Caption         =   "Format:"
         Height          =   195
         Left            =   120
         TabIndex        =   22
         Top             =   1800
         Width           =   675
      End
      Begin VB.Label lbSample3 
         AutoSize        =   -1  'True
         Caption         =   "Listening {1} by {0} From album {2}"
         Height          =   195
         Left            =   2520
         TabIndex        =   13
         Top             =   1080
         Width           =   3225
      End
      Begin VB.Label lbSample2 
         AutoSize        =   -1  'True
         Caption         =   "{1}By {0}"
         Height          =   195
         Left            =   2520
         TabIndex        =   12
         Top             =   840
         Width           =   915
      End
      Begin VB.Label lbSample1 
         AutoSize        =   -1  'True
         Caption         =   "Listen {0} - {1}"
         Height          =   195
         Left            =   2520
         TabIndex        =   11
         Top             =   600
         Width           =   1380
      End
      Begin VB.Label lbSample 
         AutoSize        =   -1  'True
         BackColor       =   &H000000F0&
         BackStyle       =   0  'Transparent
         Caption         =   "Samples"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   2640
         TabIndex        =   10
         Top             =   240
         Width           =   825
      End
      Begin VB.Label lb2 
         BackStyle       =   0  'Transparent
         Caption         =   "{2}"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   255
         Left            =   960
         TabIndex        =   8
         Top             =   1080
         Width           =   375
      End
      Begin VB.Label lb1 
         BackStyle       =   0  'Transparent
         Caption         =   "{1}"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   255
         Left            =   960
         TabIndex        =   7
         Top             =   840
         Width           =   375
      End
      Begin VB.Label lb0 
         BackStyle       =   0  'Transparent
         Caption         =   "{0}"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   255
         Left            =   960
         TabIndex        =   6
         Top             =   600
         Width           =   375
      End
      Begin VB.Label lbAlbum 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Album:"
         Height          =   195
         Left            =   240
         TabIndex        =   5
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label lbTitulo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Title:"
         Height          =   195
         Left            =   240
         TabIndex        =   4
         Top             =   840
         Width           =   435
      End
      Begin VB.Label lbArtista 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Artist:"
         Height          =   195
         Left            =   240
         TabIndex        =   3
         Top             =   600
         Width           =   525
      End
      Begin VB.Label lbLeyenda 
         AutoSize        =   -1  'True
         Caption         =   "Legend"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   360
         TabIndex        =   2
         Top             =   240
         Width           =   705
      End
      Begin VB.Shape shLeyenda 
         BackColor       =   &H0080FFFF&
         BackStyle       =   1  'Opaque
         Height          =   975
         Left            =   120
         Top             =   480
         Width           =   1575
      End
   End
   Begin VB.Image ImgWMedia 
      Height          =   240
      Left            =   7920
      Picture         =   "frmPlugin.frx":136B
      Top             =   405
      Width           =   240
   End
   Begin VB.Image ImgLogoWinamp 
      Height          =   240
      Left            =   4080
      Picture         =   "frmPlugin.frx":1537
      Top             =   405
      Width           =   240
   End
   Begin VB.Label lbWhatIHearing 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "What I'm Listening To:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2880
      TabIndex        =   0
      Top             =   120
      Width           =   2190
   End
   Begin VB.Image ImgLogo 
      Height          =   7500
      Left            =   -120
      Picture         =   "frmPlugin.frx":18C1
      Top             =   0
      Width           =   3000
   End
   Begin VB.Menu SysMenu 
      Caption         =   "SysTray Menu"
      Visible         =   0   'False
      Begin VB.Menu mnuShow 
         Caption         =   "Mostrar"
      End
      Begin VB.Menu sep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCerrar 
         Caption         =   "Cerrar"
      End
   End
End
Attribute VB_Name = "frmPlugin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'ProyPlugin: Let's Put in your Msn Messenger Windows the song
'What you're current playing in winamp
'Developed By : Int_21
'(c)2005
'Don't Forget Vote by me  ;)
'Beta version, coming soon , more features, like, multilanguage

Option Explicit
Dim sFormato$, bCustom As Boolean
Dim aInfo() As String

Private Sub Form_Load()
    AddToTray Me, Me.Caption, Me.Icon
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Dim Message As Long, temp&
   On Error Resume Next
    Message = X / Screen.TwipsPerPixelX
    Select Case Message
        Case WM_LBUTTONUP 'WM_RBUTTONUP
            'Something useful I just found out:
            ' You need to verify the height, otherwise
            ' it'll pop up the menu mid-form, if the
            ' form is big enough
            temp = GetY
            If temp > (Screen.Height / Screen.TwipsPerPixelY) - 30 Then
                PopupMenu SysMenu
            End If
    End Select
End Sub

Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then
        Me.Hide
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    NoDisplaySong
    RemoveFromTray
End Sub


' eg: Call SetMusicInfo("artist", "title", "album")
' eg: Call SetMusicInfo("artist", "title", "album", "WMContentID")
' eg: Call SetMusicInfo("artist", "title", "album", , "{1} by {0}")
' eg: Call SetMusicInfo("", "", "", , , False)
Public Sub SetMusicInfo(ByRef r_sArtist As String, ByRef r_sAlbum As String, ByRef r_sTitle As String, Optional ByRef r_sWMContentID As String = vbNullString, Optional ByRef r_sFormat As String = "{0} - {1}", Optional ByRef r_bShow As Boolean = True)

   Dim udtData As COPYDATASTRUCT
   Dim sBuffer As String
   Dim hMSGRUI As Long
   
   'Total length can not be longer then 256 characters!
   'Any longer will simply be ignored by Messenger.
   sBuffer = "\0Music\0" & Abs(r_bShow) & "\0" & r_sFormat & "\0" & r_sArtist & "\0" & r_sTitle & "\0" & r_sAlbum & "\0" & r_sWMContentID & "\0" & vbNullChar
   
   udtData.dwData = &H547
   udtData.lpData = StrPtr(sBuffer)
   udtData.cbData = LenB(sBuffer)
   
   Do
       hMSGRUI = FindWindowEx(0&, hMSGRUI, "MsnMsgrUIManager", vbNullString)
       
       If (hMSGRUI > 0) Then
           Call SendMessage(hMSGRUI, WM_COPYDATA, 0, VarPtr(udtData))
       End If
       
   Loop Until (hMSGRUI = 0)

End Sub

Private Sub lbActivar_Click()
    If lbActivar.Tag = 0 Then
        tmCheck.Enabled = True
        lbActivar.Tag = 1
        lbActivar.Caption = "Desactivar"
        DisplaySong sFormato
    Else
        lbActivar.Tag = 0
        tmCheck.Enabled = False
        lbActivar.Caption = "Activar"
        NoDisplaySong
    End If
End Sub

Private Sub lbAplicar_Click()
    DisplaySong sFormato
    lbAplicar.Enabled = False
End Sub

Private Sub lbCerrar_Click()
    Unload Me
End Sub

Private Sub mnuCerrar_Click()
    Unload Me
End Sub

Private Sub mnuShow_Click()
    Me.WindowState = vbNormal
    Me.Show
End Sub

Private Sub tmCheck_Timer()
    DisplaySong
End Sub

Function DisplaySong(Optional Formato As String)
Dim sSong$, aSong() As String
Dim sArtista$, sTema$
    If IsWinampActive Then
        sSong = GetSongName
        aSong = Split(sSong, "-")
        If UBound(aSong) > 0 Then
            sArtista = aSong(0)
            sTema = aSong(1)
        Else
            sArtista = "Winamp"
            sTema = sSong
        End If
        
        
        sFormato = txtFormato.Text
        If sFormato = "" Then sFormato = "{0}-{1}"
        Call SetMusicInfo(sArtista, "", sTema, "", sFormato)
    End If
End Function

Sub NoDisplaySong()
    Call SetMusicInfo("", "", "", "", 0, False)
End Sub

'Sub MakeStandar()
'Dim aTama%
'    aTama = 0
'    sFormato = ""
'    Erase aInfo
'    If chkArtista Then
'        aTama = aTama + 1
'        ReDim aInfo(1 To aTama)
'        aInfo(1) = "{0}"
'    End If
'    If chkTitulo Then
'        aTama = aTama + 1
'        ReDim Preserve aInfo(1 To aTama)
'        aInfo(aTama) = "{1}"
'    End If
'    If chkAlbum Then
'        aTama = aTama + 1
'        ReDim Preserve aInfo(1 To aTama)
'        aInfo(aTama) = "{2}"
'    End If
'    sFormato = Join(aInfo, "-")
'    Debug.Print sFormato
'End Sub

Private Sub txtFormato_Change()
    lbAplicar.Enabled = True
End Sub
