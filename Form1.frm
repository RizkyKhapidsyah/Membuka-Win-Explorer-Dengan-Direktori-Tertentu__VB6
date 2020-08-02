VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Membuka Win Explorer dengan Direktori Tertentu"
   ClientHeight    =   4110
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7665
   LinkTopic       =   "Form1"
   ScaleHeight     =   4110
   ScaleWidth      =   7665
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   2880
      TabIndex        =   0
      Top             =   2640
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub OpenExplorer(Optional InitialDirectory As String)
   ShellExecute 0, "Explore", InitialDirectory, vbNullString, vbNullString, SW_SHOWNORMAL
End Sub

Private Sub Command1_Click()
   'Tentukan nama direktori yang akan Anda buka dengan
   'windows explorer
   OpenExplorer ("C:\Program Files\")
End Sub


