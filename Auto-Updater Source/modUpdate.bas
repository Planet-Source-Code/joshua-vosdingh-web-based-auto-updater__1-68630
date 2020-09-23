Attribute VB_Name = "modUpdate"
'TODO: Change this to the path to your update files on a webserver.
Public Const Base = "http://www.google.com/"
Public Const MaxFiles = 300

Type File
    Name As String 'Name of file
    Version As Long 'Version as Number
    sVersion As String 'Version As String
End Type

Public Files(MaxFiles) As File 'Client Side Files
Public ServFile(MaxFiles) As File 'Server Side Files

Public Sub DownloadFile(ByVal iDl As Inet, ByVal Source As String, ByVal Target As String)
Dim Size As Long
Dim Size2 As String
Dim Remaining As Long
Dim dl() As Byte
Dim f As Byte

    f = FreeFile 'Free File

    iDl.Execute Base & Source, "GET" 'Start the Download
    Do While iDl.StillExecuting 'Make sure we dont download when its not ready to.
        DoEvents
    Loop

    Size2 = iDl.GetHeader("Content-Length") 'Set the size
    Size = CLng(Size2)
    Remaining = Size
    
    frmUpdate.lblStatus.Caption = "Downloading " & Source & " ..." 'Update our label, so users know whats going on.
    Open App.Path & "\" & Target For Binary As #f 'Open the target file
        Do While Remaining > 0 'Keep getting chunks of data untill theres none left.
            DoEvents 'Give gui chance to update
            If Not frmUpdate.Visible Then frmUpdate.Show 'Make sure the Form is shown.
            If Remaining > 1024 Then 'More than a KB Left
                dl = iDl.GetChunk(1024, icByteArray) 'Download the chunk.
                Remaining = Remaining - 1024 'Set our new remaining byte count
            Else 'Less than a KB Left
                dl = iDl.GetChunk(Remaining, icByteArray) 'Download the rest of the data
                Remaining = 0 'Set it to 0 so we know to break out of the loop
            End If
            DoEvents 'Give gui change to do stuff again.
            frmUpdate.barStatus.Value = (Size - Remaining) / Size * 100 'Update the Progress Bar
            Put #f, , dl() 'Put the file contents into the target file
        Loop
    Close #f
End Sub

'Used for finding a client side file out of the list
Public Function FindFile(ByVal Name As String) As Integer
    FindFile = -1
    For I = 0 To MaxFiles
        If Files(I).Name = Name Then FindFile = I
    Next I
End Function

'Used for finding a server side file out of the list
Public Function FindServ(ByVal Name As String) As Integer
    FindServ = -1
    For I = 0 To MaxFiles
        If ServFile(I).Name = Name Then FindServ = I
    Next I
End Function
