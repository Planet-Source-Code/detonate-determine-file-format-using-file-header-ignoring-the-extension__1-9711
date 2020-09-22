<div align="center">

## Determine file format using file header, ignoring the extension


</div>

### Description

Determines the filetype (PE Executable, GIF Image, Word Document, etc) using the file header/contents, as opposed to using the file extension. Supported types: PE Executable, INI File, AVI Movie, WAV Audio, Word document, Access database, GIF Image, MP3 audio, BMP image, TIFF image, ZIP archive, ARJ archive, RAR archive, HTML/HTA docs, JPEG image, Visual Basic files. If it can't determine the filetype, itll try and determine if the file is text or binary.
 
### More Info
 
Public Function GetFileType(xFile As String)

'xFile is the full path:\filename of the file to test.

Returns a string indicating the filetype, or if the file is an unknown text or unknown binary file.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Detonate](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/detonate.md)
**Level**          |Beginner
**User Rating**    |4.4 (22 globes from 5 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0, VB Script, ASP \(Active Server Pages\) 
**Category**       |[Files/ File Controls/ Input/ Output](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/files-file-controls-input-output__1-3.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/detonate-determine-file-format-using-file-header-ignoring-the-extension__1-9711/archive/master.zip)





### Source Code

```
Public Function GetFileType(xFile As String) As String
On Error Resume Next
Dim ID As String * 300
If Dir$(xFile) = "" Then
  GetFileType = "NOT FOUND"
  Exit Function
End If
Open xFile For Binary Access Read As #1
 Get #1, 1, ID
Close #1
If Left(ID, 2) = "MZ" Or Left(ID, 2) = "ZM" Then
  GetFileType = "PE Executable"
  Exit Function
ElseIf Left(ID, 1) = "[" And InStr(1, Left(ID, 100), "]") > 0 Then
  GetFileType = "INI File"
  Exit Function
ElseIf Mid(ID, 9, 8) = "AVI LIST" Then
  GetFileType = "AVI Movie File"
  Exit Function
ElseIf Left(ID, 4) = "RIFF" Then
  GetFileType = "WAV Audio File"
  Exit Function
ElseIf Left(ID, 4) = Chr(208) & Chr(207) & Chr(17) & Chr(224) Then
  GetFileType = "Microsoft Word Document"
  Exit Function
ElseIf Mid(ID, 5, 15) = "Standard Jet DB" Then
  GetFileType = "Microsoft Access Database"
  Exit Function
ElseIf Left(ID, 3) = "GIF" Or InStr(1, ID, "GIF89") > 0 Then
  GetFileType = "GIF Image File"
  Exit Function
ElseIf Left(ID, 1) = Chr(255) And Mid(ID, 5, 1) = Chr(0) Then
  GetFileType = "MP3 Audio File"
  Exit Function
ElseIf Left(ID, 2) = "BM" Then
  GetFileType = "BMP (Bitmap) Image File"
  Exit Function
ElseIf Left(ID, 3) = "II*" Then
  GetFileType = "TIFF Image File"
  Exit Function
ElseIf Left(ID, 2) = "PK" Then
  GetFileType = "ZIP Archive File"
  Exit Function
ElseIf InStr(1, LCase(ID), "<html>") > 0 Or InStr(1, LCase(ID), "<!doctype") > 0 Then
  GetFileType = "HTML Document File"
  Exit Function
ElseIf UCase(Left(ID, 3)) = "RAR" Then
  GetFileType = "RAR Archive File"
  Exit Function
ElseIf Left(ID, 2) = Chr(96) & Chr(234) Then
  GetFileType = "ARJ Archive File"
  Exit Function
ElseIf Left(ID, 3) = Chr(255) & Chr(216) & Chr(255) Then
  GetFileType = "JPEG Image File"
  Exit Function
ElseIf InStr(1, ID, "Type=") > 0 And InStr(1, ID, "Reference=") > 0 Then
  GetFileType = "Visual Basic Project File"
  Exit Function
ElseIf Left(ID, 8) = "VBGROUP " Then
  GetFileType = "Visual Basic Group Project File"
  Exit Function
ElseIf Left(ID, 8) = "VERSION " & InStr(1, ID, vbCrLf & "Begin") > 0 Then
  GetFileType = "Visual Basic Form File"
  Exit Function
Else
 'Unknown file... make a weak attempt to determine if the file is text or binary
 If InStr(1, ID, Chr$(255)) > 0 Or InStr(1, ID, Chr$(1)) > 0 Or InStr(1, ID, Chr$(2)) > 0 Or InStr(1, ID, Chr$(3)) > 0 Then
  GetFileType = "Unknown binary file"
 Else
  GetFileType = "Unknown text file"
 End If
 Exit Function
End If
End Function
```

