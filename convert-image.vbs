
' settings:
Const IM_PATH       = "C:\Program Files\ImageMagick-7.0.2-Q16-HDRI\"  ' the path to ImageMagick
Const IM_IO_BUFFER_SIZE = 4096                          ' number of characters in IO buffer



' some more constants needed
Const IM_CONVERT    = "convert.exe"                     ' the convert tool
Const IM_COMPOSITE  = "composite.exe"                   ' the composite tool
Const IM_MOGRIFY    = "mogrify.exe"                     ' the mogrify tool



''
' This function opens a "Select Folder" dialog and will
' return the fully qualified path of the selected folder
'
' Argument:
'     folder           [string]    the root folder where you can start browsing;
'                                  if an empty string is used, browsing starts
'                                  on the local computer
'
'     prompt           [string]    the message to be shown to the user
'
' Returns:
' A string containing the fully qualified path of the selected folder
'
' Written by Rob van der Woude
' http://www.robvanderwoude.com
Function SelectFolder( ByVal folder, ByVal prompt )

    ' Standard housekeeping
    Dim objShell, objFolder
    
    ' Custom error handling
    On Error Resume Next
    SelectFolder = vbNull

    ' Create a dialog object
    Set objShell  = CreateObject("Shell.Application")
    Set objFolder = objShell.BrowseForFolder(0, prompt, 0, folder)

    ' Return the path of the selected folder
    If IsObject( objfolder ) Then SelectFolder = objFolder.Self.Path

    ' Standard housekeeping
    Set objFolder = Nothing
    Set objshell  = Nothing
    On Error Goto 0
    
End Function



''
' This function runs a command in background and returns it's exit code.
' Argument:
'     command           [string]  a command which can contain environment variables
'     output            [string&] a reference to a variable that will receive the command's output
'
' Returns:
' A numeric value indicating the command's exit code.
Function RunCommand( ByVal command, ByRef output )
  
  Dim objShell, objExec, strOutput
  
  On Error Resume Next
  Set objShell = CreateObject("WScript.Shell")
  Set objExec = objShell.Exec(command)
  strOutput = ""
  
  Do While Not objExec.StdOut.AtEndOfStream
    strOutput = strOutput & objExec.StdOut.Read(IM_IO_BUFFER_SIZE)
  Loop

  Do While Not objExec.StdErr.AtEndOfStream
    strOutput = strOutput & objExec.StdErr.Read(IM_IO_BUFFER_SIZE)
  Loop

  Do While objExec.Status = 0
     WScript.Sleep 100
  Loop

  output = strOutput
  RunCommand = objExec.ExitCode
  
  Set objExec = Nothing
  Set objShell = Nothing
  On Error Goto 0
  
End Function



''
' Queries multiple elements by their ids.
'
Function GetElementsByIds ( ByVal ids(), ByRef result() )
  
  Dim i

  ReDim result(UBound(ids))
  
  For i = 0 To UBound(ids) - 1
    Set result(i) = document.getElementById(ids(i))
  Next

  GetElementsById = result

End Function



''
' Returns the argument that specifies the stamp size.
'
Function GetStampSizeArgument ( )
  
  Dim stampWidth, stampHeight

  stampWidth = document.getElementById("stampWidth").value
  stampHeight = document.getElementById("stampHeight").value

  GetStampSizeArgument = "-size " & CStr(stampWidth) & "x" & CStr(stampHeight)

End Function



''
' Returns the argument that specifies the stamp's font.
'
Function GetStampFontArgument ( )
  
  Dim stampFont

  stampFont = document.getElementById("stampFont").value

  GetStampFontArgument = "-font """ & stampFont & """"

End Function



''
' Returns the argument that specifies the stamp font size.
'
Function GetStampFontSizeArgument ( )
  
  Dim stampFontSize

  stampFontSize = document.getElementById("stampFontSize").value

  GetStampFontSizeArgument = "-pointsize " & CStr(stampFontSize)

End Function



''
' Returns the argument that specifies the maximum size of the processed images.
'
Function GetResizeArgument ( )
  
  Dim resizeImageCheck, maxImageWidth, maxImageHeight

  GetResizeArgument = ""
  resizeImageCheck = document.getElementById("resizeImageCheck").checked
  maxImageWidth = document.getElementById("maxImageWidth").value
  maxImageHeight = document.getElementById("maxImageHeight").value

  If resizeImageCheck Then
    GetResizeArgument = "-filter Cubic -resize " & CStr(maxImageWidth) & "x" & CStr(maxImageHeight) & "^>"
  End If

End Function



''
' Returns the selected file extension for the source file.
'
Function GetSrcExtension ( )
  
  Dim i, ids(4), srcExtension(), checked

  ids(0) = "srcExtension1"
  ids(1) = "srcExtension2"
  ids(2) = "srcExtension3"
  ids(3) = "srcExtension4"

  GetElementsByIds ids, srcExtension
  GetSrcExtension = "jpg"

  For i = 0 To UBound(srcExtension) - 1 
    checked = srcExtension(i).checked
    If CBool(checked) Then
      GetSrcExtension = CStr(srcExtension(i).value)
    End If
  Next

End Function



''
' Returns the selected file extension for the destination file.
'
Function GetDstExtension ( )
  
  Dim i, ids(4), dstExtension(), checked

  ids(0) = "dstExtension1"
  ids(1) = "dstExtension2"
  ids(2) = "dstExtension3"
  ids(3) = "dstExtension4"

  GetElementsByIds ids, dstExtension
  GetDstExtension = "jpg"

  For i = 0 To UBound(dstExtension) - 1
    checked = dstExtension(i).checked
    If CBool(checked) Then
      GetDstExtension = CStr(dstExtension(i).value)
    End If
  Next

End Function



''
' Returns the selected stamp position.
'
Function GetStampPos ( )
  
  Dim i, ids(9), stampPos(), checked

  ids(0) = "stampPos1"
  ids(1) = "stampPos2"
  ids(2) = "stampPos3"
  ids(3) = "stampPos4"
  ids(4) = "stampPos5"
  ids(5) = "stampPos6"
  ids(6) = "stampPos7"
  ids(7) = "stampPos8"
  ids(8) = "stampPos9"

  GetElementsByIds ids, stampPos
  GetStampPos = "southwest"

  For i = 0 To UBound(stampPos) - 1
    checked = stampPos(i).checked
    If CBool(checked) Then
      GetStampPos = CStr(stampPos(i).value)
    End If
  Next

End Function



''
' Calculates the stamp filename.
'
Function GetStampFilename ( )
  
  Dim fso, dstFolder

  Set fso = CreateObject("Scripting.FileSystemObject")
  dstFolder = document.getElementById("dstFolderTxt").innerHTML
  GetStampFilename = fso.BuildPath(dstFolder, "stamp.png")

  Set fso = Nothing

End Function



''
' Returns the source folder.
'
Function GetSrcFolder ( )
  
  GetSrcFolder = document.getElementById("srcFolderTxt").innerHTML

End Function



''
' Returns all files in the source directory in respect with the source file extension.
'
Function GetSrcFiles ( )

  Dim fso, folder, files, file, result(), i

  Set fso = CreateObject("Scripting.FileSystemObject")
  Set folder = fso.GetFolder(GetSrcFolder())
  Set files = folder.Files

  i = 0
  ReDim Preserve result(0)
  For Each file In files
    If fso.GetExtensionName(LCase(file.Name)) = LCase(GetSrcExtension()) Then
      ReDim Preserve result(i + 1)
      result(i) = file.Path
      i = i + 1
    End If
  Next

  GetSrcFiles = result

  Set files = Nothing
  set folder = Nothing
  Set fso = Nothing

End Function



''
' Returns the destination directory.
'
Function GetDstFolder ( )
  
  GetDstFolder = document.getElementById("dstFolderTxt").innerHTML

End Function



''
' Returns a boolean value indicating if stamp image should be trimmed.
'
Function GetStampTrim ( )
  
  GetStampTrim = document.getElementById("stampTrim").checked

End Function



''
' Start batch converting images.
'
Function ConvertImages ( )
  
  Dim fso, command, output, srcFiles, i, srcPath, dstFolder, dstExtension, dstPath

  Set fso = CreateObject("Scripting.FileSystemObject")
  ConvertImages = False
  srcFiles = GetSrcFiles()
  dstFolder = GetDstFolder()
  dstExtension = GetDstExtension()
  SetStatus ""

  ' for each image filename
  For i = 0 To UBound(srcFiles) - 1
    
    srcPath = srcFiles(i)
    dstPath = fso.BuildPath(dstFolder, fso.GetBaseName(srcPath) & "." & dstExtension)
    command = """" & IM_PATH & IM_COMPOSITE & """ " & _
              "-gravity " & GetStampPos() & " -geometry +10+10 """ & GetStampFilename() & """ " & GetResizeArgument() & " """ & srcPath & """ """ & dstPath & """"

    AppendStatus command & Chr(13)

    If RunCommand(command, output) <> 0 Then
      AppendStatus Chr(13) & output
      Exit Function
    End If

  Next

  ConvertImages = True
  Set fso = Nothing

End Function



''
' Create a stamp file.
'
' composite -compose CopyOpacity stamp_bg.png stamp_fg.png stamp.png
Function CreateStamp ( )
  
  Dim fso, extension, filename, fgFilename, bgFilename

  CreateStamp = False
  Set fso = CreateObject("Scripting.FileSystemObject")

  filename = GetStampFilename()
  extension = fso.GetExtensionName(filename)
  fgFilename = fso.BuildPath(fso.GetParentFolderName(filename), fso.GetBaseName(filename) & "_fg." & extension)
  bgFilename = fso.BuildPath(fso.GetParentFolderName(filename), fso.GetBaseName(filename) & "_bg." & extension)

  If CreateStampFg(fgFilename) And CreateStampBg(bgFilename) Then
    
    Dim command, output
    command = """" & IM_PATH & IM_COMPOSITE & """ " & _ 
              "-compose CopyOpacity """ & bgFilename & """ """ & fgFilename & """ """ & filename & """"

    SetStatus command
    If RunCommand(command, output) = 0 Then
      
      If GetStampTrim() Then
        command = """" & IM_PATH & IM_MOGRIFY & """ " & _
                  "-trim +repage """ & filename & """"

        SetStatus command

        If RunCommand(command, output) = 0 Then
          CreateStamp = True
        Else
          AppendStatus Chr(13) & output
        End If

      Else
        CreateStamp = True
      End If

    Else
      AppendStatus Chr(13) & output
    End If
    
  End If

  Set fso = Nothing

End Function



''
' Create the stamp file foreground.
'
' convert -size 300x50 xc:grey30 -font Arial -pointsize 20 -gravity center -draw "fill grey70  text 0,0  'Copyright'" stamp_fg.png
Function CreateStampFg ( ByVal filename )
  
  Dim stampText, command, output

  CreateStampFg = False
  stampText = document.getElementById("stampTxt").value
  command = """" & IM_PATH & IM_CONVERT & """ " & _
            GetStampSizeArgument () & " xc:grey30 " & GetStampFontArgument() & " " & GetStampFontSizeArgument () & " -gravity center -draw ""fill grey70  text 0,0  '" & stampText & _
            "'"" """ & filename & """"
  
  SetStatus command
  If RunCommand(command, output) = 0 Then
    CreateStampFg = True
  Else
    AppendStatus Chr(13) & output
  End If

End Function



''
' Create the stamp file background.
'
' convert -size 300x50 xc:black -font Arial -pointsize 20 -gravity center -draw "fill white  text  1,1  'Copyright' text  0,0  'Copyright' fill black  text -1,-1 'Copyright'" +matte stamp_bg.png
Function CreateStampBg ( ByVal filename )
  
  Dim stampText, command, output

  CreateStampBg = False
  stampText = document.getElementById("stampTxt").value
  command = """" & IM_PATH & IM_CONVERT & """ " & _
            GetStampSizeArgument () & " xc:black " & GetStampFontArgument() & " " & GetStampFontSizeArgument () & " -gravity center " & _
            "-draw """ & _
            "fill white  text  1,1  '" & stampText & "' " & _
            "            text  0,0  '" & stampText & "' " & _
            "fill black  text -1,-1 '" & stampText & "' " & _
            "'"" +matte """ & filename & """"
  
  SetStatus command
  If RunCommand(command, output) = 0 Then
    CreateStampBg = True
  Else
    AppendStatus Chr(13) & output
  End If

End Function



''
' Handle OnClick event of browseSrcFolderBtn.
'
Sub browseSrcFolderBtn_OnClick ( evt )
        
  Dim srcFolderTxt, folder
  
  evt.returnValue = false
  Set srcFolderTxt = document.getElementById("srcFolderTxt")

  folder = SelectFolder("", "Verzeichnis mit Quellbildern auswählen")
        
  If Not folder = vbNull Then
          
    While srcFolderTxt.hasChildNodes()
      srcFolderTxt.removeChild(srcFolderTxt.firstChild)
    Wend
          
    srcFolderTxt.appendChild(document.createTextNode(folder))
          
  End If

  Set srcFolderTxt = Nothing

End Sub



''
' Handle OnClick event of browseDstFolderBtn.
'
Sub browseDstFolderBtn_OnClick ( evt )
        
  Dim dstFolderTxt, folder
        
  evt.returnValue = false
  Set dstFolderTxt = document.getElementById("dstFolderTxt")

  folder = SelectFolder("", "Zielverzeichnis für Bilder auswählen")
        
  If Not folder = vbNull Then
          
    While dstFolderTxt.hasChildNodes()
      dstFolderTxt.removeChild(dstFolderTxt.firstChild)
    Wend
          
    dstFolderTxt.appendChild(document.createTextNode(folder))
          
  End If

  Set dstFolderTxt = Nothing

End Sub



''
' Handle OnClick event of startConversionBtn.
'
Sub startConversionBtn_OnClick ( evt )
        
  evt.returnValue = false
        
  ' Wasserzeichen wird immer als PNG erstellt (Transparenz)
  If CreateStamp() Then
    SetStatus "Wasserzeichen erfolgreich erstellt."
          
    ' Batchkonvertierung starten...
    If ConvertImages() Then
      SetStatus "Bilder erfolgreich konvertiert."
    Else
      AppendStatus "Fehler bei der Konvertierung."
    End If
          
  Else
    AppendStatus "Konnte Wasserzeichen nicht erstellen."
  End If
        
End Sub



''
' Handle OnClick event of resizeImageCheck.
'
Sub resizeImageCheck_OnClick ( evt )
        
  Dim checked, maxImageWidth, maxImageHeight
        
  checked = document.getElementById("resizeImageCheck").checked
  Set maxImageWidth = document.getElementById("maxImageWidth")
  Set maxImageHeight = document.getElementById("maxImageHeight")
        
  If checked Then
    maxImageWidth.disabled = False
    maxImageHeight.disabled = False
  Else
    maxImageWidth.disabled = True
    maxImageHeight.disabled = True
  End If

  Set maxImageWidth = Nothing
  Set maxImageHeight = Nothing

End Sub



''
' Clears the status panel.
'
Sub ClearStatus ( )
        
  Dim statusTxt
  Set statusTxt = document.getElementById("statusTxt")
        
  While statusTxt.hasChildNodes()
    statusTxt.removeChild(statusTxt.firstChild)
  Wend
  
  Set statusTxt = Nothing

End Sub



''
' Sets the status panel content.
'
Sub SetStatus ( ByVal text )
        
  Dim statusTxt
  Set statusTxt = document.getElementById("statusTxt")
        
  While statusTxt.hasChildNodes()
    statusTxt.removeChild(statusTxt.firstChild)
  Wend
        
  statusTxt.appendChild(document.createTextNode(text))
        
  Set statusTxt = Nothing

End Sub



''
' Appends status panel content.
'
Sub AppendStatus ( ByVal text )
        
  Dim statusTxt
  Set statusTxt = document.getElementById("statusTxt")
        
  statusTxt.appendChild(document.createTextNode(text))
        
  Set statusTxt = Nothing

End Sub



''
' Startup routine.
'
Sub Main ( )
        
  Dim browseSrcFolderBtn, browseDstFolderBtn, startConversionBtn, output, version
  Dim resizeImageCheck, fso, cwd
  
  Set fso = CreateObject("Scripting.FileSystemObject")
  cwd = fso.GetAbsolutePathName(".")
        
  document.getElementById("srcFolderTxt").appendChild(document.createTextNode(cwd))
  document.getElementById("dstFolderTxt").appendChild(document.createTextNode(cwd))

  Set browseSrcFolderBtn = document.getElementById("browseSrcFolderBtn")
  Set browseSrcFolderBtn.onclick = GetRef("browseSrcFolderBtn_OnClick")

  Set browseDstFolderBtn = document.getElementById("browseDstFolderBtn")
  Set browseDstFolderBtn.onclick = GetRef("browseDstFolderBtn_OnClick")

  Set startConversionBtn = document.getElementById("startConversionBtn")
  Set startConversionBtn.onclick = GetRef("startConversionBtn_OnClick")
       
  Set resizeImageCheck = document.getElementById("resizeImageCheck")
  Set resizeImageCheck.onclick = GetRef("resizeImageCheck_OnClick")
        
  Set version = document.getElementById("version")
  If RunCommand("""" & IM_PATH & IM_CONVERT & """" & " -version", output) = 0 Then
    version.appendChild(document.createTextNode(output))
  Else
    version.appendChild(document.createTextNode("Es wurde ein Problem mit der ImageMagick Version geunden." & Chr(13)))
    version.appendChild(document.createTextNode(output))
  End If
  
  Set version = Nothing
  Set resizeImageCheck = Nothing
  Set startConversionBtn = Nothing
  Set browseDstFolderBtn = Nothing
  Set browseSrcFolderBtn = Nothing
  Set fso = Nothing

End Sub