
dZips = Now - 1

sFolder = "C:\intelitrader\log\InteliOrder"
Set oFSO = CreateObject("Scripting.FileSystemObject")

For Each oFile In oFSO.GetFolder(sFolder).Files

    strFileType = Right(oFile.Name,3)
    
    if strFileType = "log" then	
		  iYear = Left(oFile.Name, 4)
		  iMonth = Mid(oFile.Name, 6, 2)
      iDay = Mid(oFile.Name, 9, 2)

      d = CDate(iYear & "/" & iMonth & "/" & iDay)

      if d < dZips then
        WindowsZip sFolder & "\" & oFile.Name, sFolder & "\" & iYear & "." & iMonth & "." & iDay & ".zip"
      end if

	end if

Next


Function WindowsZip(sFile, sZipFile)

  Set oZipShell = CreateObject("WScript.Shell") 
  Set oZipFSO = CreateObject("Scripting.FileSystemObject")

  If Not oZipFSO.FileExists(sZipFile) Then
    NewZip(sZipFile)
  End If

  Set oZipApp = CreateObject("Shell.Application")
  sZipFileCount = oZipApp.NameSpace(sZipFile).items.Count
  aFileName = Split(sFile, "\")
  sFileName = (aFileName(Ubound(aFileName)))

  'listfiles
  sDupe = False

  For Each sFileNameInZip In oZipApp.NameSpace(sZipFile).items
    If LCase(sFileName) = LCase(sFileNameInZip) Then
      sDupe = True
      Exit For
    End If

  Next
 
  If Not sDupe Then
    oZipApp.NameSpace(sZipFile).Copyhere sFile
    'Keep script waiting until Compressing is done
    On Error Resume Next
    sLoop = 0
    Do Until sZipFileCount < oZipApp.NameSpace(sZipFile).Items.Count
      Wscript.Sleep(100)
      sLoop = sLoop + 1
    Loop
    On Error GoTo 0
  End If
  
  Wscript.Sleep(100)
  oFSO.DeleteFile (sFile)

End Function


Sub NewZip(sNewZip)
  Set oNewZipFSO = CreateObject("Scripting.FileSystemObject")
  Set oNewZipFile = oNewZipFSO.CreateTextFile(sNewZip)

  oNewZipFile.Write Chr(80) & Chr(75) & Chr(5) & Chr(6) & String(18, 0)
  oNewZipFile.Close

  Set oNewZipFSO = Nothing
  Wscript.Sleep(500)
End Sub