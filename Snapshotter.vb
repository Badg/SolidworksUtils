' ******************************************************************************
' LICENSING
' -------------------------------------------------
' 
' Solidworks Snapshotter: Generate .jpg + .step (part/assy) or .pdf (dwg) of file
'     Copyright (C) 2014-2015 Nicholas Badger
'     badg@nickbadger.com
'     nickbadger.com
' 
'     This library is free software; you can redistribute it and/or
'     modify it under the terms of the GNU Lesser General Public
'     License as published by the Free Software Foundation; either
'     version 2.1 of the License, or (at your option) any later version.
' 
'     This library is distributed in the hope that it will be useful,
'     but WITHOUT ANY WARRANTY; without even the implied warranty of
'     MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
'     Lesser General Public License for more details.
' 
'     You should have received a copy of the GNU Lesser General Public
'     License along with this library; if not, write to the Free Software
'     Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA  02110-1301
'     USA
' 
' ------------------------------------------------------
' ******************************************************************************
Dim swApp           As SldWorks.SldWorks
Dim swFile          As SldWorks.ModelDoc2
Dim swFile2         As SldWorks.ModelDocExtension
Dim swFileType      As Integer
Dim filePath        As String
Dim filePathBare    As String
Dim fileDir         As String
Dim fileName        As String
Dim snapFolder      As String
Dim snapBaseName    As String
Dim dateString      As String

Dim successString   As String
Dim failString      As String

Dim boolstatus      As Boolean
Dim longstatus      As Long, longwarnings As Long

Sub main()

    Set swApp = Application.SldWorks
    Set swFile = swApp.ActiveDoc
    Set swFile2 = swFile.Extension
    snapFolder = "Snapshots"

    If Not swFile Is Nothing Then
        swFileType = swFile.GetType
        filePath = swFile.GetPathName
        ' Strip the extension
        filePathBare = Left(filePath, InStrRev(filePath, ".") - 1)
        ' Get the filename
        fileName = Right(filePathBare, Len(filePathBare) - InStrRev(filePathBare, "\"))
        ' Get the directory
        fileDir = Left(filePathBare, InStrRev(filePathBare, "\"))
        ' Construct the current date string
        dateString = CStr(Year(Now)) + " " + CStr(Format(Month(Now), "00")) + " " + CStr(Format(Day(Now), "00"))
        ' Get the extensionless name of the files to create in snapshots directory
        snapBaseName = fileDir + snapFolder + "\" + fileName + " - " + dateString

        ' Verify file folder exists, or create it (quick and dirty)
        If Len(Dir(fileDir + snapFolder, vbDirectory)) = 0 Then
            MkDir (fileDir + snapFolder)
        End If
    Else
        End
    End If

    If swFileType = swDocPART Or swFileType = swDocASSEMBLY Then
        longstatus = swFile.SaveAs3(snapBaseName + ".jpg", 0, 0)
        longstatus = swFile.SaveAs3(snapBaseName + ".step", 0, 0)

    ElseIf swFileType = swDocDRAWING Then
        longstatus = swFile.SaveAs3(snapBaseName + ".pdf", 0, 0)
    End If

    MsgBox "Success?"

End Sub