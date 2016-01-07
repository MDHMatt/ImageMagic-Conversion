'************************************************************
'********* PDF to PNG ImageMagick Conversion Script *********
'********* Author: Matthew Dinsdale *************************
'********* Url: http://mattdinsdale.uk **********************
'********* Date: 05/1/2016 **********************************
'************************************************************
'********* Notes:-    ***************************************
' This script will require ImageMagick which can
' be found at http://www.imagemagick.org/
' During installation ensure that ImageMagickObject OLE for
' VBscript, Visual Basic and WSH is checked. Development
' Headers for C and C++ optional.
'
' This script uses ImageMagick to convert PFD's to PNG but
' can be used for other formats. Just change the FromExt &
' ToExt variables as required.
'
' Define StartFolder in the function to limit the searchable
' folder. If left as "" the whole local computer and network
' is usable. 
'************************************************************

option Explicit
' Variable Names
Dim Path
Dim InputFile
Dim OutputFile
Dim img
Dim InputFolder
Dim fsoFolder
Dim NameLenght
Dim fso
Dim FromExt
Dim ToExt

' Set File extension to convert From to New extension ie .PDF to .PNG ** The leading . (dot) is required! **
FromExt = ".pdf"
ToExt = ".png"

' Returned From Function
Path = SelectFolder( "" )
' MsgBox Path

If Path = vbNull Then
    MsgBox "Conversion Cancelled",64,"Cancelled"
    Else

    ' Define Variable Functions
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set fsoFolder = fso.GetFolder(Path)
    Set img = CreateObject("ImageMagickObject.MagickImage.1")
    Set InputFolder = fsoFolder.Files
    
    For Each InputFile in InputFolder
        If Right(InputFile,4) = FromExt Then
            NameLenght = Len(InputFile) - 4
            OutputFile = Mid(InputFile,1,NameLenght)
            OutputFile = OutputFile & ToExt
            
            ' **** For Debugging Comment out if not needed ****
            
            'MsgBox "Input: " & InputFile.Path
            'MsgBox "Input: " & OutputFile
            
            ' ImageMagick Conversion
            img.convert InputFile.Path, OutputFile
        End If
    Next

MsgBox "PDF to PNG Conversion Complete",64,"Complete"

End If

Function SelectFolder( StartFolder )
 ' This function opens a "Select Folder" dialog and will
 ' return the fully qualified path of the selected folder
 '
 ' Argument:
 '     StartFolder    [string]      the root folder where you can start browsing;
 '                                  if an empty string is used, browsing starts
 '                                  on the local computer
 '
 ' Returns:
 ' A string containing the fully qualified path of the selected folder
 '
 ' Written by Rob van der Woude http://www.robvanderwoude.com
 ' Adjusted by Matt Dinsdale http://mattdinsdale.uk
 
    ' Standard housekeeping
    Dim objFolder, objItem, objShell
    
    'Define Starting folder
    StartFolder = "\\cadserver\users\eng\dinsdale_m\"
    
    ' Custom error handling
    On Error Resume Next
    SelectFolder = vbNull

    ' Create a dialog object
    Set objShell  = CreateObject( "Shell.Application" )
    Set objFolder = objShell.BrowseForFolder( 0, "Select folder to convert PDF's to PNG", 0, StartFolder )

    ' Return the path of the selected folder
    If IsObject( objfolder ) Then SelectFolder = objFolder.Self.Path
    
    Set objShell.CurrentDirectory = SelectFolder
    
    ' Standard housekeeping
    Set objFolder = Nothing
    Set objshell  = Nothing
    On Error Goto 0
End Function
