VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Compression/Decompression"
   ClientHeight    =   1890
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1890
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog dlgCommon 
      Left            =   0
      Top             =   1440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdCompress 
      Caption         =   "Compress File"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1560
      Width           =   1575
   End
   Begin VB.CommandButton cmdDecompress 
      Caption         =   "Decompress File"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3000
      TabIndex        =   6
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Frame Frame2 
      Caption         =   "File Out"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   4455
      Begin VB.TextBox txtOut 
         Height          =   285
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   3615
      End
      Begin VB.CommandButton cmdOut 
         Caption         =   "..."
         Height          =   285
         Left            =   3840
         TabIndex        =   4
         Top             =   240
         Width           =   435
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "File In"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      Begin VB.CommandButton cmdIn 
         Caption         =   "..."
         Height          =   285
         Left            =   3840
         TabIndex        =   2
         Top             =   240
         Width           =   435
      End
      Begin VB.TextBox txtIn 
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   3615
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdDecompress_Click()
    Compression_DeCompress txtIn.Text, txtOut.txt, LZW
End Sub

Private Sub cmdIn_Click()
    dlgCommon.FileName = App.Path
    dlgCommon.ShowOpen
    txtIn.Text = dlgCommon.FileName
End Sub

Private Sub cmdOut_Click()
    dlgCommon.FileName = App.Path
    dlgCommon.ShowSave
    txtOut.Text = dlgCommon.FileName
End Sub

Private Sub cmdCompress_Click()
    Compression_Compress txtIn.Text, txtOut.txt, LZW
End Sub
