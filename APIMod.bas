Attribute VB_Name = "APIMod"

Option Explicit

Public blnExit As Boolean


Private Declare Function lstrcat Lib "kernel32" _
                Alias "lstrcatA" (ByVal lpString1 As String, ByVal _
                lpString2 As String) As Long


' Declarations and such needed for the example:
' (Copy them to the (declarations) section of a module.)
Public Declare Function GetFileVersionInfo Lib "version.dll" Alias "GetFileVersionInfoA" _
              (ByVal lptstrFilename As String, ByVal dwHandle As Long, ByVal dwLen As Long, _
              lpData As Any) As Long
Public Declare Function GetFileVersionInfoSize Lib "version.dll" Alias _
              "GetFileVersionInfoSizeA" (ByVal lptstrFilename As String, lpdwHandle As Long) As Long
Public Declare Function VerQueryValue Lib "version.dll" Alias "VerQueryValueA" (pBlock _
              As Any, ByVal lpSubBlock As String, lplpBuffer As Long, puLen As Long) As Long
Public Declare Function lstrcpy Lib "kernel32.dll" Alias "lstrcpyA" (ByVal lpString1 _
              As Any, ByVal lpString2 As Any) As Long
'Public Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, _
              Source As Any, ByVal Length As Long)
Public Type VS_FIXEDFILEINFO
    dwSignature As Long
    dwStrucVersion As Long
    dwFileVersionMS As Long
    dwFileVersionLS As Long
    dwProductVersionMS As Long
    dwProductVersionLS As Long
    dwFileFlagsMask As Long
    dwFileFlags As Long
    dwFileOS As Long
    dwFileType As Long
    dwFileSubtype As Long
    dwFileDateMS As Long
    dwFileDateLS As Long
End Type



Public Sub Get_File_Version(sDirName As String)

  Dim i As Integer
  Dim lngFileVerInfo As Long
  Dim sFileName As String
  Dim hWndOwner As Long
  Dim count As Integer
  'Dim lpData ':( As Variant ?
  Dim vffi As VS_FIXEDFILEINFO  ' version info structure
  Dim buffer() As Byte          ' buffer for version info resource
  Dim pData As Long             ' pointer to version info data
  Dim nDataLen As Long          ' length of info pointed at by pData
  Dim cpl(0 To 3) As Byte       ' buffer for code page & language
  Dim cplstr As String          ' 8-digit hex string of cpl
  Dim dispstr As String         ' string used to display version information
  Dim retval As Long            ' generic return value
  Dim sExtension As String
  Dim ItemX As ListItem
  Dim x As Integer
  
    blnExit = False
    Form1.lvwFileList.ListItems.Clear
    Form1.MousePointer = vbHourglass
    For i = 0 To Form1.File1.ListCount
        DoEvents
        If blnExit = True Then
            Form1.MousePointer = vbNormal
            Exit Sub
        End If
        sExtension = Right$(Form1.File1.List(i), 3)
        If UCase$(sExtension) = "DLL" Or UCase$(sExtension) = "EXE" Or UCase$(sExtension) = "OCX" Then
           
            sFileName = sDirName + "\" + Form1.File1.List(i)
            If sFileName <> "" Then
                ' First, get the size of the version info resource.  If this function fails, then Text1
                ' identifies a file that isn't a 32-bit executable/DLL/etc.
                nDataLen = GetFileVersionInfoSize(sFileName, pData)
                If nDataLen <> 0 Then
                    ' Make the buffer large enough to hold the version info resource.
                    ReDim buffer(0 To nDataLen - 1) As Byte
                    ' Get the version information resource.
                    retval = GetFileVersionInfo(sFileName, 0, nDataLen, buffer(0))
    
                    ' Get a pointer to a structure that holds a bunch of data.
                    retval = VerQueryValue(buffer(0), "\", pData, nDataLen)
                    ' Copy that structure into the one we can access.
                    CopyMemory vffi, ByVal pData, nDataLen
                    ' Display the full version number of the file.
                    dispstr = Trim$(Str$(HIWORD(vffi.dwFileVersionMS))) & "." & _
                              Trim$(Str$(LOWORD(vffi.dwFileVersionMS))) & "." & _
                              Trim$(Str$(HIWORD(vffi.dwFileVersionLS))) & "." & _
                              Trim$(Str$(LOWORD(vffi.dwFileVersionLS)))
                    Debug.Print "Version Number: "; dispstr
                    
                    Set ItemX = Form1.lvwFileList.ListItems.Add(, , Form1.File1.List(i))
                    ItemX.SubItems(1) = dispstr
                    
                End If
           
              Else
                Exit Sub
            End If
        End If
    
    Next i
Form1.MousePointer = vbNormal
Form1.mnuFileExit.Enabled = True
End Sub

' *** Place the following function definitions inside a module. ***

' HIWORD and LOWORD are API macros defined below.
Public Function HIWORD(ByVal dwValue As Long) As Long

  Dim hexstr As String

    hexstr = Right$("00000000" & Hex$(dwValue), 8)
    HIWORD = CLng("&H" & Left$(hexstr, 4))

End Function

Public Function LOWORD(ByVal dwValue As Long) As Long

  Dim hexstr As String

    hexstr = Right$("00000000" & Hex$(dwValue), 8)
    LOWORD = CLng("&H" & Right$(hexstr, 4))

End Function

' This nifty subroutine swaps two byte values without needing a buffer variable.
' This technique, which uses Xor, works as long as the two values to be swapped are
' numeric and of the same data type (here, both Byte).
Public Sub SwapByte(byte1 As Byte, byte2 As Byte)

    byte1 = byte1 Xor byte2
    byte2 = byte1 Xor byte2
    byte1 = byte1 Xor byte2

End Sub

' This function creates a hexadecimal string to represent a number, but it
' outputs a string of a fixed number of digits.  Extra zeros are added to make
' the string the proper length.  The "&H" prefix is not put into the string.
Public Function FixedHex(ByVal hexval As Long, ByVal nDigits As Long) As String

    FixedHex = Right$("00000000" & Hex$(hexval), nDigits)

End Function

':) Ulli's Code Formatter V2.0 (6/26/01 9:34:01 AM) 68 + 134 = 202 Lines
