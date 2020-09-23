VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "msinet.ocx"
Begin VB.Form frmUpdate 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Auto-Updater"
   ClientHeight    =   1095
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4455
   Icon            =   "frmUpdate.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1095
   ScaleWidth      =   4455
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.StatusBar staBar 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   840
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   3889
            Text            =   "Current Version is Unknown"
            TextSave        =   "Current Version is Unknown"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   3889
            Text            =   "Latest Version is Unknown"
            TextSave        =   "Latest Version is Unknown"
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraStatus 
      Caption         =   "Auto Updater Status"
      Height          =   855
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4455
      Begin MSComctlLib.ProgressBar barStatus 
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label lblStatus 
         Caption         =   "Status Unknown..."
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
         TabIndex        =   1
         Top             =   240
         Width           =   4215
      End
   End
   Begin InetCtlsObjects.Inet Inet 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
End
Attribute VB_Name = "frmUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Version As Long 'Version as number
Public sVersion As String 'Version as string
Public updDone As Boolean 'For file protection
Public Latest As Long 'Latest version as number
Public sLatest As String 'lastest version as string

Private Sub Form_Load()
On Error GoTo Err
Dim f As Byte
Dim Temp As String
Dim Temp2() As String
Dim First As Boolean
Dim Line As Long

    f = FreeFile
    Load frmUpdate 'Make sure the form is shown.
    frmUpdate.Show
    updDone = True 'Used for closing the window, True = You can close, False = You can't
               'Stops corrupt files when downloading/Decompressing

    lblStatus.Caption = "Downloading News..." 'Update Caption
    Call DownloadFile(Inet, "News.txt", "News.txt") 'Download the current news
    f = FreeFile 'Free File
    Open App.Path & "\News.txt" For Binary As #f 'Get the news now
        Temp = Input(LOF(f), #f)
    Close #f

    'Update the news form and show it.
    frmNews.txtNews.Text = "Close this window to continue." & vbNewLine & "===========================================" & vbNewLine & vbNewLine & Temp
    frmNews.Top = Me.Top + Me.Height + 25
    frmNews.Left = Me.Left + Me.Width / 2 - frmNews.Width / 2
    frmNews.Show

    lblStatus.Caption = "Getting Local Information ..." 'Update caption
    If Dir(App.Path & "\v.dat") = "" Then 'If v.dat doesnt exist create it.
        Open App.Path & "\v.dat" For Binary As #f
            Put #f, , "0.0.0"
        Close #f
        f = FreeFile
    End If

    Open App.Path & "\v.dat" For Binary As #f 'Load the local data
        Do While Not EOF(f)
            Line Input #f, Temp
            If Left(Temp, 1) <> ";" Then 'Look, i added comments. Live with it :p
                If Not First Then 'First line is program version
                    sVersion = Temp
                    Version = CLng(Replace(sVersion, ".", "")) 'remove dots and save as number
                    First = True 'Set it so we know we passed the first line
                Else 'Other lines are individual file versions
                    Temp2 = Split(Temp, ",") 'Get the data, its seporated by commas
                    Files(Line).Name = Temp2(0) 'First is file name/path
                    Files(Line).sVersion = Temp2(1) 'second is file version
                    Files(Line).Version = CLng(Replace(Files(Line).sVersion, ".", "")) 'Save version as number, same as above.
                    Line = Line + 1 'So we know where to add the data in the array.
                End If
            End If
        Loop
    Close #f
    staBar.Panels(1).Text = "Current Version is " & sVersion 'Tell them what version they have

    Call DownloadFile(Inet, "v.dat", "s.dat") 'Download the Remote Version File
    lblStatus.Caption = "Getting Remote Information ..." 'Update... Well you know
    f = FreeFile
    First = False
    Line = 0
    
                    'all this is the same as above, except with Remote File list/ServFiles Array
    Open App.Path & "\s.dat" For Binary As #f
        Do While Not EOF(f)
            Line Input #f, Temp
            If Left(Temp, 1) <> ";" Then
                If Not First Then
                    sLatest = Temp
                    Latest = CLng(Replace(sLatest, ".", ""))
                    First = True
                Else
                    Temp2 = Split(Temp, ",")
                    ServFile(Line).Name = Temp2(0)
                    ServFile(Line).sVersion = Temp2(1)
                    ServFile(Line).Version = CLng(Replace(ServFile(Line).sVersion, ".", ""))
                    Line = Line + 1
                End If
            End If
        Loop
    Close #f
    staBar.Panels(2).Text = "Latest Version is " & sLatest 'Show them latest version

    If Latest > Version Then 'They need to update, ask them if they want to.
        'TODO: Change Program Name
        If MsgBox("There is a update avaliable to update PROGRAM NAME from V" & sVersion & " To V" & sLatest & vbNewLine & "Update PROGRAM NAME now?", vbYesNo + vbInformation + vbApplicationModal, "Update Avaliable") = vbNo Then
            Kill App.Path & "\s.dat" 'They don't want to update.. Retards...
            Kill App.Path & "\News.txt"
            End
        End If
    Else 'Good for them, they have the latest version
        'TODO: Change program Name
        Call MsgBox("Your currently using the latest version of the Soul Society Online! There is no need to update it.", vbInformation + vbOKOnly, "No Update Avaliable")
        Kill App.Path & "\s.dat"
        Kill App.Path & "\News.txt"
        End
    End If

    updDone = False 'They cant exit the update now cz were dealing with important files!
    For I = 0 To MaxFiles 'Update the files
        If ServFile(I).Name <> "" Then 'Make sure there is actully a file here.
            If FindFile(ServFile(I).Name) <> -1 Then 'See if we have the file already
                If Files(FindFile(ServFile(I).Name)).Version < ServFile(I).Version Then 'We do, check to see if it needs updating.
                    Call DownloadFile(Inet, ServFile(I).Name, ServFile(I).Name & ".tmp") 'Download the compressed file to a temp file
                    lblStatus.Caption = "Decompressing " & ServFile(I).Name & " ..." ' Show the caption.
                    DoEvents 'Update stuff
                    Compression_DeCompress App.Path & "\" & ServFile(I).Name & ".tmp", App.Path & "\" & ServFile(I).Name, LZW 'Decompress the file
                    Kill App.Path & "\" & ServFile(I).Name & ".tmp" 'Bye Bye Temp Compressed file!
                End If
            Else 'They dont have the file, download it.
                Call DownloadFile(Inet, ServFile(I).Name, ServFile(I).Name & ".tmp") 'Downloading
                lblStatus.Caption = "Decompressing " & ServFile(I).Name & " ..." 'Caption
                DoEvents 'Update
                Compression_DeCompress App.Path & "\" & ServFile(I).Name & ".tmp", App.Path & "\" & ServFile(I).Name, LZW 'Decompress
                Kill App.Path & "\" & ServFile(I).Name & ".tmp" 'Beheading :p
                'Same deal as above.
            End If
        End If
    Next I 'Do for every file untill there up to date.

    'Update everything
    Kill App.Path & "\s.dat"
    Kill App.Path & "\News.txt"
    f = FreeFile
    barStatus.Value = 0
    lblStatus.Caption = "Updating Info..."

    Open App.Path & "\v.dat" For Binary As #f 'Make a new Client Side File List
        Put #f, , sLatest
        For I = 0 To MaxFiles
           Line = FindFile(ServFile(I).Name) 'Find client side file
            If ServFile(I).Name <> "" Then 'Make sure its a file
                If Line < 0 Then 'Do they have the file?
                    'Nope
                    Put #f, , vbNewLine & ServFile(I).Name & "," & ServFile(I).sVersion
                ElseIf Files(Line).Version <= ServFile(I).Version Then 'They have it, only replace version info
                                                                       'If its a older file, (might be newer cz of beta editions etc.)
                    Put #f, , vbNewLine & ServFile(I).Name & "," & ServFile(I).sVersion
                End If
            End If
            If Files(I).Name <> "" And FindServ(Files(I).Name) < 0 Then 'Client side files server doesnt have, dont remove them.
                Put #f, , vbNewLine & Files(I).Name & "," & Files(I).sVersion
            End If
            barStatus.Value = I / MaxFiles * barStatus.Max
        Next I
    Close #f
    lblStatus.Caption = "Update Finished. You can now close this window." 'Yay! Were done!!!
    staBar.Panels(1).Text = "Current Version is " & sLatest 'Update current caption to latest.
    updDone = True 'Were done

Err:
    If Err.Number > 0 Then 'Zomg there was an error! Wth....
        On Error Resume Next
        updDone = True
        'TODO: Put your email here.
        Call MsgBox("There was an error updating! Please send the following information to <email here>" & vbNewLine & "Error #" & Err.Number & " - " & Err.Description & "." & vbNewLine & "This data has been copied to your ClipBoard", vbCritical, "Error Updating")
        Clipboard.Clear
        Clipboard.SetText "Error #" & Err.Number & " - " & Err.Description & "." 'Set there clipboard
        Kill App.Path & "\s.dat" 'Kill the not needed files
        Kill App.Path & "\News.txt"
        End
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If updDone = True Then 'They can exit
        Kill App.Path & "\s.dat" 'Kill the not needed files
        Kill App.Path & "\News.txt"
        End
    Else 'They cant exit. Warn them
        Call MsgBox("Sorry, but you can't exit now." & vbNewLine & "Some files may be Fragmented causing the program to become corrupt." & vbNewLine & "Please wait untill the update finishes to exit.", vbCritical, "Update in progress...")
        Cancel = 1
    End If
End Sub
