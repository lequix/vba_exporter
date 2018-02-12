'------------------------------------------------------------
'VBA File Exporter
'------------------------------------------------------------
'Auther: Hironori Yamanoha
'Ver: 2.0
'Date: 2018-02-06
'Description: This scripts allows to export VBS files from.
'             an Excel file.
'-------------------------------------------------------------

Dim fso
Dim fl 'File instance
Dim xlsm_fnd 'Boolean
Dim xlsm_file 'Filename
Dim target_dir 'Target Directory
Dim lib_dir 'Lib Directory (Place for common cls files.)
Dim obj_Excel, obj_workBook 'Object for opening an Excel file.
Dim obj_param 'args
Dim temp_comp

'A01: Initial Setup
'---------------------
xlsm_fnd = False 'Init flag for xlsm detect.

'A02: Get a XLSM File in the current directory
'---------------------
Set fso= CreateObject("Scripting.FileSystemObject")

'Set target directory
Set obj_param = WScript.Arguments
if obj_param.Count > 0 then
	'Parameter(s) detected
	target_dir = fso.getParentFolderName(WScript.ScriptFullName) & "\" & obj_param.item(0)
	lib_dir = fso.getParentFolderName(WScript.ScriptFullName) 
else
	'No parametors detected
	'Target direcotory is set as the parent folder
	target_dir = fso.getParentFolderName(WScript.ScriptFullName)
	lib_dir = Left(target_dir, InStrRev(target_dir, "\", -1, vbBinaryCompare))
end if

'Check Excel/Macro file under the target directory specified.
For Each fl In fso.GetFolder(target_dir).Files
	If Right(fl.name, 4) = "xlsm" then
		If xlsm_fnd = False then 
			xlsm_fnd = True
			xlsm_file = fl.name
		Else
			'Exit this script because more than one XLM file
			'was found.
			MsgBox(xlsm_file & ": More than one XLSM file was found in the directory.")
			Set fso = Nothing 'Release FSO
			WScript.Quit
		End If
	End If
Next
Set fso = Nothing 'Release fso

Set obj_Excel = CreateObject("Excel.Application")
obj_Excel.Visible = False
obj_Excel.DisplayAlerts = False
obj_Excel.EnableEvents = False

'A03: Open the Excel/Macro file
'---------------------
Set obj_workBook = obj_Excel.Workbooks.Open(target_dir & "\" & xlsm_file)

'A04: Export Module
Call ExportSource()

obj_Excel.DisplayAlerts = True
obj_Excel.EnableEvents = True
obj_workBook.Close False
obj_Excel.Quit
Set obj_workBook = nothing
Set obj_Excel = nothing

'M01: Export Source
'---------------------
Sub ExportSource()
    For Each temp_comp In obj_workBook.VBProject.VBComponents
	If temp_comp.CodeModule.CountOfDeclarationLines <> temp_comp.CodeModule.CountofLines Then
            Select Case temp_comp.Type
                'STANDARD_MODULE
                Case 1
										If Left(temp_comp.Name, 4) = "Mod_" Then
                    	temp_comp.Export target_dir & "\" & temp_comp.Name & ".bas"
										Else
                    	temp_comp.Export target_dir & "\" & obj_workBook.Name & "_" & temp_comp.Name & ".bas"
										End If
                'CLASS_MODULE
                Case 2
										temp_comp.Export
                    'temp_comp.Export target_dir & "\" & obj_workBook.Name & "_" & temp_comp.Name & ".cls"
                    temp_comp.Export target_dir & "\" & temp_comp.Name & ".cls"
                'USER_FORM
                Case 3
                    temp_comp.Export target_dir & "\" & obj_workBook.Name & "_" & temp_comp.Name & ".frm"
                'SHEET?ThisWorkBook
                Case 100
                    temp_comp.Export target_dir & "\" & obj_workBook.Name & "_" & temp_comp.Name & ".bas"
            End Select
            With temp_comp.CodeModule
                'Code = .Lines(1, .CountOfLines)
                'Code = .Lines(.CountOfDeclarationLines + 1, .CountOfLines - .CountOfDeclarationLines + 1)                    
            End With
	End If
    Next
End Sub
