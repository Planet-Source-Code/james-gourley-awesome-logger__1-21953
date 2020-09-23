Attribute VB_Name = "Module1"
Sub DestroyFile(sFileName As String)
    On Error Resume Next
    Dim Block1 As String, Block2 As String, Blocks As Long
    Dim hFileHandle As Integer, iLoop As Long, offset As Long
    'Create two buffers with a specified 'wi
    '     pe-out' characters
    Const BLOCKSIZE = 4096
    Block1 = String(BLOCKSIZE, "X")
    Block2 = String(BLOCKSIZE, " ")
    'Overwrite the file contents with the wi
    '     pe-out characters
    hFileHandle = FreeFile
    Open sFileName For Binary As hFileHandle
    Blocks = (LOF(hFileHandle) \ BLOCKSIZE) + 1


    For iLoop = 1 To Blocks
        offset = Seek(hFileHandle)
        Put hFileHandle, , Block1
        Put hFileHandle, offset, Block2
    Next iLoop
    Close hFileHandle
    'Now you can delete the file, which cont
    '     ains no sensitive data
    Kill sFileName
End Sub
