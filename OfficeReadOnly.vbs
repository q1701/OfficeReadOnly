' OfficeReadOnly
'
' Copyright (c) 2013 q1701
' Licensed under the MIT license: http://opensource.org/licenses/MIT
'
' Description:
'   Open Microsoft Office files as read only.
'   Supported file types are below.
'    (doc / xls / ppt / mpp / vsd)
' Usage:
'   Drag files and Drop them onto this script.

Option Explicit

'Entry Point
Dim officeReadOnlyAppl
Set officeReadOnlyAppl = New OfficeReadOnly
officeReadOnlyAppl.Main WScript.Arguments
Set officeReadOnlyAppl = Nothing
WScript.Quit

'Main application class
Class OfficeReadOnly
    'Variables
    Private fsObj           'FileSystemObject
    Private shObj           'Shell
    Private mWordObj        'Word
    Private mExcelObj       'Excel
    Private mPowerPointObj  'Power Point
    Private mProjectObj     'Project
    Private mVisioObj       'Visio
    'Initializer
    Private Sub Class_Initialize
        Set fsObj = WScript.CreateObject("Scripting.FileSystemObject")
        Set shObj = WScript.CreateObject("WScript.Shell")
        Set mWordObj = Nothing
        Set mExcelObj = Nothing
        Set mPowerPointObj = Nothing
        Set mProjectObj = Nothing
        Set mVisioObj = Nothing
    End Sub
    'Terminator
    Private Sub Class_Terminate
        'Release objects
        Set fsObj = Nothing
        Set shObj = Nothing
        Set mWordObj = Nothing
        Set mExcelObj = Nothing
        Set mPowerPointObj = Nothing
        Set mProjectObj = Nothing
        Set mVisioObj = Nothing
    End Sub
    'Private Utilities
    Private Function CreateOrGetObject(progId)
        'Create or Get object
        Dim obj
        On Error Resume Next
        Set obj = GetObject(, progId)
        'No active instance found
        '(ActiveX component can't create object)
        If Err.Number = 429 Then
            On Error Goto 0
            Set obj = CreateObject(progId)
        End If
        Set CreateOrGetObject = obj
    End Function
    'Properties
    Public Property Get wordObj
        If mWordObj Is Nothing Then
            Set mWordObj = CreateOrGetObject("Word.Application")
        End If
        Set wordObj = mWordObj
    End Property
    Public Property Get excelObj
        If mExcelObj Is Nothing Then
            Set mExcelObj = CreateOrGetObject("Excel.Application")
        End If
        Set excelObj = mExcelObj
    End Property
    Public Property Get powerPointObj
        If mPowerPointObj Is Nothing Then
            Set mPowerPointObj = CreateOrGetObject("PowerPoint.Application")
        End If
        Set powerPointObj = mPowerPointObj
    End Property
    Public Property Get projectObj
        If mProjectObj Is Nothing Then
            Set mProjectObj = CreateOrGetObject("Msproject.Application")
        End If
        Set projectObj = mProjectObj
    End Property
    Public Property Get visioObj
        If mVisioObj Is Nothing Then
            Set mVisioObj = CreateOrGetObject("Visio.Application")
        End If
        Set visioObj = mVisioObj
    End Property
    'Open procedure's
    Private Sub OpenWithWord(fullFileName)
        wordObj.Visible = True
        wordObj.Documents.Open fullFileName, , True
        shObj.AppActivate  wordObj.Caption
    End Sub
    Private Sub OpenWithExcel(fullFileName)
        excelObj.Visible = True
        excelObj.Workbooks.Open fullFileName, , True
        shObj.AppActivate  excelObj.Caption
    End Sub
    Private Sub OpenWithPowerPoint(fullFileName)
        powerPointObj.Visible = True
        powerPointObj.Presentations.Open fullFileName, True
        shObj.AppActivate  powerPointObj.Caption
    End Sub
    Private Sub OpenWithProject(fullFileName)
        projectObj.Visible = True
        projectObj.FileOpen fullFileName, True
        shObj.AppActivate  projectObj.Caption
    End Sub
    Private Sub OpenWithVisio(fullFileName)
        visioObj.Documents.OpenEx fullFileName, visOpenRO
        shObj.AppActivate  visioObj.Caption
    End Sub
    'Usage
    Private Sub Usage
        WScript.Echo(   "OfficeReadOnly"                                    & vbCrLf _
                    &   "  Description:"                                    & vbCrLf _
                    &   "    Open Microsoft Office files as read only."     & vbCrLf _
                    &   "    Supported file types are below."               & vbCrLf _
                    &   "     (doc / xls / ppt / mpp / vsd)"                & vbCrLf _
                    &   "  Usage:"                                          & vbCrLf _
                    &   "    Drag files and Drop them onto this script.")
    End Sub
    '------
    ' Main
    '------
    Public Sub Main(ByRef Arguments)
        'Variables
        Dim fullFileName    'fullpath
        Dim fileName        'filename
        Dim fileExt         'filename extension
        'Show usage if no files are given.
        If Arguments.Unnamed.Count = 0 Then
            Usage
            Exit Sub
        End If
        'Process each files.
        For Each fullFileName In Arguments
            'Extract the filename extension
            fileName = fsObj.GetFileName(fullFileName)
            fileExt = LCase(fsObj.GetExtensionName(fullFileName))
            'Determine the file type
            Select Case fileExt
            Case "doc", "docx", "docm"
                OpenWithWord fullFileName
            Case "xls", "xlsx", "xlsm"
                OpenWithExcel fullFileName
            Case "ppt", "pptx", "pptm"
                OpenWithPowerPoint fullFileName
            Case "vsd"
                OpenWithVisio fullFileName
            Case "mpp"
                OpenWithProject fullFileName
            Case Else
                'Skip unsupported files
            End Select
        Next
    End Sub
End Class
