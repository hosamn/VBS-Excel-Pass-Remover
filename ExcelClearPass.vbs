' #===============  by: Hosam El Nagar ===============#

Set objExcel = CreateObject("Excel.Application")

objExcel.Visible = FALSE
objExcel.DisplayAlerts = TRUE
crrnt_pas = "es1959"
neew_pass = ""

' MsgBox "Merge" & "Text"
' scriptdir = """" & scriptdir & """"


Set oFSO = CreateObject("Scripting.FileSystemObject")

scriptdir = oFSO.GetParentFolderName(WScript.ScriptFullName)

For Each f in oFSO.GetFolder(scriptdir).Files
  If InStr(f, "xlsx") Then
    If InStr(f, "~") Then
    Else
      ' MsgBox f

      file_path = f
      save_path = f & "-unlocked.xlsx"
      ' MsgBox save_path

      ' expression.Open (FileName, UpdateLinks, ReadOnly, Format, Password, WriteResPassword, IgnoreReadOnlyRecommended, Origin, Delimiter, Editable, Notify, Converter, AddToMru, Local, CorruptLoad)
      Set objWorkbook = objExcel.Workbooks.Open(file_path,,,,crrnt_pas)

      objWorkbook.Password = neew_pass
      objWorkbook.SaveAs save_path

      objExcel.Quit
    
    END IF
  End If
Next



