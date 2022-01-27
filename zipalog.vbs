'dtmOntem = Date() - 15
'dtmHoje = Date()

dNow = Now

'Declarations
Dim objFSO, strPath
Const SHCONTF_NONFOLDERS = &H40
strPath = "C:\intelitrader\log\InteliOrder\"  
Set objFSO = CreateObject("Scripting.FileSystemObject")

Scanfiles (objFSO.GetFolder(strPath)) 

Set objFSO=Nothing
wscript.quit

'Sub ScanDirectory(objFolder)
'    Scanfiles objFolder
'    For Each fld In objFolder.SubFolders 
'      ScanDirectory fld 
'    Next
'End Sub

Sub Scanfiles(objFolder)
    Wscript.Echo objFolder
    For Each fil In objFolder.files 
        Wscript.Echo fil
		'Compress (fil) 
    Next

'Set objFSOFILE = CreateObject("Scripting.FileSystemObject")
'objFSOFILE.MoveFile "C:\log\*.zip" , "c:\teste2\"  'AQUI VOCÊ COLOCA O CAMINHO QUE QUER MOVER OS ZIP.
End Sub

Sub Compress(fil)
  strPath = Left(fil, InStrRev(fil, "\"))     
  strFile = Mid(fil, InStrRev(fil, "\") + 1)     
  strExt = Mid(strFile, InStrRev(strFile, ".") + 1, 3)     
  If LCase(strExt) = "log" Then           
    strZip = strPath & Replace(strFile, "log", "zip")
    set zipFil=objFSO.CreateTextFile(strZip)
    zipFil.WriteLine Chr(80) & Chr(75) & Chr(5) & Chr(6) & String(18, 0)
    zipfil.Close
   Set oApp = CreateObject("Shell.Application")
   oApp.NameSpace(strZip).CopyHere strPath & strFile
   wscript.sleep 500 'Meio segundo para comprimir o arquivo - dependendo do tamanho, aumente este valor
   set oApp=Nothing
   Set zipFil=Nothing
   If objFSO.FileExists(strZip) Then objFSO.DeleteFile (fil)
   End If

End sub