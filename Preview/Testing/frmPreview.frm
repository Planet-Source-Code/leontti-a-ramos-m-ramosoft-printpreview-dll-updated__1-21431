VERSION 5.00
Begin VB.Form frmPreview 
   Caption         =   "Form1"
   ClientHeight    =   6300
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9465
   LinkTopic       =   "Form1"
   ScaleHeight     =   6300
   ScaleWidth      =   9465
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "&Test"
      Height          =   255
      Left            =   2520
      TabIndex        =   2
      Top             =   0
      Width           =   915
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Host Preview"
      Height          =   255
      Left            =   180
      TabIndex        =   1
      Top             =   0
      Width           =   1695
   End
   Begin VB.PictureBox picHost 
      Height          =   5895
      Left            =   0
      ScaleHeight     =   5835
      ScaleWidth      =   9315
      TabIndex        =   0
      Top             =   300
      Width           =   9375
   End
End
Attribute VB_Name = "frmPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim m_oPreview As Preview

Private Sub Check1_Click()
    m_oPreview.Container = (picHost.hWnd * Check1.Value)
End Sub

Private Sub Command1_Click()
    With m_oPreview
        .Cls
        With .Pages
            .ScaleMode = vbInches
            .Width = 8.5
            .Height = 11
            .Add
            With .ActivePage
                .DrawPicture 3.25, 0.5, 2, 1.2, LoadPicture(App.Path & "\RamoSoft.gif"), True
                .SetFont "Tahoma", 24, True
                .DrawText "RamoSoft Print Preview Dll", 1, 1.5, 6, 2, vbBlue, , vbCenter
                .SetFont "Tahoma", 14, , True
                .DrawText "This demo show drawing capabilities of the RamoSoft Print Preview dll", _
                0.8, 3, 7, 2, vbBlack, vbCyan, vbCenter
                .SetFont "OCR A Extended", 72, True, , , , -45
                .DrawText "Cool!", 2, 3, 5, 4, vbRed
                '.DrawBox 0.5, 0.5, 1, 2.5, vbBlue, vbRed, 2
            End With
        End With
        .Show
    End With
End Sub

Private Sub Form_Load()
    Set m_oPreview = New Preview
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Set m_oPreview = Nothing
End Sub


