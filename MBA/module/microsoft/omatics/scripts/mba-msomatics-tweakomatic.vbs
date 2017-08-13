
Dim strDBFile

Const adOpenStatic = 3
    Const adLockOptimistic = 3
    Const adUseClient = 3
    DefaultComputer = "."
    MasterFile = ""
    RetrievalFile = ""
    StartHelp = "To begin, select a manageable component, and then select a category of tasks. When you do so, a set of tasks will be displayed in the Task Area. Click a task and two scripts will automatically be created: one for configuring information, the other for retrieving information."
    Help2 = "Select a category from the list of categories. When you do so, a set of tasks will be displayed in the Task Area. Click a task and two scripts will automatically be created: one for configuring information, the other for retrieving information."
    Help3 = "Click a task and two scripts will automatically be created: one for configuring information, the other for retrieving information."


Sub Window_Onload_Tweakomatic
	strDBFile = MBA.mdl_cur.PathDB & "\mba-msomatics-tweakomatic.mdb" '"module\microsoft\msomatics\data\db\mba-msomatics-tweakomatic.mdb"
    Set objConnection = CreateObject("ADODB.Connection") 
    objConnection.Open "Provider= Microsoft.Jet.OLEDB.4.0; Data Source=" & strDBFile
    Set objRecordset = CreateObject("ADODB.Recordset")
	
    objRecordset.CursorLocation = adUseClient
    objRecordset.Open "SELECT DISTINCT Tweaks.Component FROM Tweaks ORDER BY Tweaks.Component" , objConnection,  adOpenStatic, adLockOptimistic
    objRecordSet.MoveFirst
    strHTML = "<select style='width: 100%' onChange=""GetCategoryInfo()"" name=ComponentList>" 
    strHTML = strHTML & "<option value= " & chr(34) & chr(34) & ">"
    Do Until objRecordSet.EOF
        strHTML = strHTML & "<option value= " & chr(34) & _
            objRecordSet.Fields.Item("Component") & chr(34) & _
                ">" & objRecordSet.Fields.Item("Component")
        objRecordSet.MoveNext
    Loop
    strHTML = strHTML & "</select>"
    ComponentArea.InnerHTML = strHTML
    oFormTweakomatic.run_button.disabled = True
    oFormTweakomatic.run_button2.disabled = True
    oFormTweakomatic.save_button.disabled = True
    oFormTweakomatic.save_button2.disabled = True    
    oFormTweakomatic.change_button.disabled = True
    oFormTweakomatic.Master_button.disabled = True
    oFormTweakomatic.Master_button2.disabled = True
    oFormTweakomatic.show_button.disabled = True
    oFormTweakomatic.show_button2.disabled = True
    HelpArea.InnerHTML = StartHelp
End Sub



Sub Window_OnUnload
   On Error Resume Next
   Set objFSO = CreateObject("Scripting.FileSystemObject")
   objFSO.DeleteFile "temp_script.sm"
   Set objFSO = Nothing
   self.Close()
End Sub


Sub GetCategoryInfo()
    On Error Resume Next
    FilterValue = oFormTweakomatic.ComponentList.Value
    Set objConnection = CreateObject("ADODB.Connection") 
    objConnection.Open "Provider= Microsoft.Jet.OLEDB.4.0; Data Source=" & strDBFile
    Set objRecordset = CreateObject("ADODB.Recordset")
    objRecordset.CursorLocation = adUseClient
    objRecordset.Open "SELECT DISTINCT Tweaks.Category FROM Tweaks Where Component = '" & FilterValue & "'ORDER BY Tweaks.Category" , objConnection,  adOpenStatic, adLockOptimistic
    objRecordSet.MoveFirst
    CategoryArea.InnerHTML = ""
    strHTML = "<select style='width:100%' onChange=""GetTaskInfo()"" name=CategoryList>" 
    strHTML = strHTML & "<option value= " & chr(34) & chr(34) & ">"
    Do Until objRecordSet.EOF
        strHTML = strHTML & "<option value= " & chr(34) &_
            objRecordSet.Fields.Item("Category") & chr(34) &_
                ">" & objRecordSet.Fields.Item("Category")
        objRecordSet.MoveNext
    Loop
    strHTML = strHTML & "</select>"
    CategoryArea.InnerHTML = strHTML
    TaskArea.InnerHTML = "<select size='15' name='D2'></select>"
    HelpArea.InnerHTML= Help2
    oFormTweakomatic.ScriptArea.innerText = ""
    oFormTweakomatic.RetrievalArea.innerText = ""
    oFormTweakomatic.run_button.disabled = True
    oFormTweakomatic.run_button2.disabled = True
    oFormTweakomatic.save_button.disabled = True
    oFormTweakomatic.save_button2.disabled = True
    oFormTweakomatic.change_button.disabled = True
    oFormTweakomatic.Master_button.disabled = True
    oFormTweakomatic.Master_button2.disabled = True
    oFormTweakomatic.objRecordSet.Close
    oFormTweakomatic.objConnection.Close
End Sub


Sub GetTaskInfo()
    On Error Resume Next
    FilterValue = oFormTweakomatic.ComponentList.Value
    FilterValue2 = oFormTweakomatic.CategoryList.Value	
    Set objConnection = CreateObject("ADODB.Connection") 
    objConnection.Open "Provider= Microsoft.Jet.OLEDB.4.0; Data Source=" & strDBFile
    Set objRecordset = CreateObject("ADODB.Recordset")
    objRecordset.CursorLocation = adUseClient
    objRecordset.Open "SELECT DISTINCT Tweaks.Task FROM Tweaks Where Component = '" & FilterValue & "' AND Category = '" & FilterValue2 & "'ORDER BY Tweaks.Task" , objConnection,  adOpenStatic, adLockOptimistic
    objRecordSet.MoveFirst
    TaskArea.InnerHTML = ""
    strHTML = "<select size = '15' style='width: 100%' onChange=""GetHelpText()"" name=TaskList>" 
    Do Until objRecordSet.EOF
        strHTML = strHTML & "<option value= " & chr(34) &_
            objRecordSet.Fields.Item("Task") & chr(34) &_
                ">" & objRecordSet.Fields.Item("Task")
        objRecordSet.MoveNext
    Loop
    strHTML = strHTML & "</select>"
    TaskArea.InnerHTML = strHTML
    HelpArea.InnerHTML = Help3
    oFormTweakomatic.ScriptArea.innerText = ""
    oFormTweakomatic.RetrievalArea.innerText = ""
    oFormTweakomatic.run_button.disabled = True
    oFormTweakomatic.run_button2.disabled = True
    oFormTweakomatic.save_button.disabled = True
    oFormTweakomatic.save_button2.disabled = True
    oFormTweakomatic.change_button.disabled = True
    oFormTweakomatic.Master_button.disabled = True
    oFormTweakomatic.Master_button2.disabled = True
    objRecordSet.Close
    objConnection.Close
End Sub


Sub GetHelpText()
    FilterValue = oFormTweakomatic.TaskList.Value
    Set objConnection = CreateObject("ADODB.Connection") 
    objConnection.Open "Provider= Microsoft.Jet.OLEDB.4.0; Data Source=" & strDBFile
    Set objRecordset = CreateObject("ADODB.Recordset")
    objRecordset.CursorLocation = adUseClient
    objRecordset.Open "SELECT * FROM Tweaks WHERE Task = '" & FilterValue & "' ORDER BY Tweaks.Task" , objConnection,  adOpenStatic, adLockOptimistic
    objRecordSet.MoveFirst
    Do Until objRecordSet.EOF
        strHTML = objRecordSet.Fields.Item("Help") 
        strText2 = "On Error Resume Next" & vbCrLf
        strLocation  = objRecordSet.Fields.Item("RegistryLocation") 
        If StrLocation = "HKEY_CURRENT_USER" Then
            strText = "HKEY_CURRENT_USER = &H80000001" & VbCrLf 
            strText2 = strText2 & "HKEY_CURRENT_USER = &H80000001" & VbCrLf

            strText = strText & "strComputer = " & chr(34) & DefaultComputer & chr(34) & VbCrLf
            strText2 = strText2 & "strComputer = " & chr(34) & DefaultComputer & chr(34) & VbCrLf

            strText = strText & "Set objReg = GetObject(" & chr(34) & "winmgmts:\\" & chr(34)
            strText2 = strText2 & "Set objReg = GetObject(" & chr(34) & "winmgmts:\\" & chr(34)

            strText = strText & " & strComputer & " & chr(34) & "\root\default:StdRegProv" & chr(34) & ")" & VbCrLf
            strText2 = strText2 & " & strComputer & " & chr(34) & "\root\default:StdRegProv" & chr(34) & ")" & VbCrLf

            strText = strText & "strKeyPath = " & chr(34) & objRecordSet.Fields.Item("RegKey") & chr(34) & VbCrLf 
            strText2 = strtext2 & "strKeyPath = " & chr(34) & objRecordSet.Fields.Item("RegKey") & chr(34) & VbCrLf 

            strText = strText & "objReg.CreateKey " & strLocation & ", strKeyPath" & VbCrLf

            strText = strText & "ValueName = " & chr(34) & objRecordSet.Fields.Item("RegValue") & chr(34) & VbCrLf
            strText2 = strText2 & "ValueName = " & chr(34) & objRecordSet.Fields.Item("RegValue") & chr(34) & VbCrLf

        Else
            strText = strText & "HKEY_LOCAL_MACHINE = &H80000002" & vbCrLf
            strText2 = strText2 & "HKEY_LOCAL_MACHINE = &H80000002" & vbCrLf

            strText = strText & "strComputer = " & chr(34) & DefaultComputer & chr(34) & VbCrLf
            strText2 = strText2 & "strComputer = " & chr(34) & DefaultComputer & chr(34) & VbCrLf

            strText = strText & "Set objReg = GetObject(" & chr(34) & "winmgmts:\\" & chr(34)
            strText2 = strText2 & "Set objReg = GetObject(" & chr(34) & "winmgmts:\\" & chr(34)

            strText = strText & " & strComputer & " & chr(34) & "\root\default:StdRegProv" & chr(34) & ")" & VbCrLf
            strText2 = strText2 & " & strComputer & " & chr(34) & "\root\default:StdRegProv" & chr(34) & ")" & VbCrLf

            strText = strText & "strKeyPath = " & chr(34) & objRecordSet.Fields.Item("RegKey") & chr(34) & VbCrLf 
            strText2 = strText2 & "strKeyPath = " & chr(34) & objRecordSet.Fields.Item("RegKey") & chr(34) & VbCrLf

            strText = strText & "objReg.CreateKey " & strLocation & ", strKeyPath" & VbCrLf

            strText = strText & "ValueName = " & chr(34) & objRecordSet.Fields.Item("RegValue") & chr(34) & VbCrLf
            strText2 = strText2 & "ValueName = " & chr(34) & objRecordSet.Fields.Item("RegValue") & chr(34) & VbCrLf

        End If

        strValueType = objRecordSet.Fields.Item("DataType")

        strText2 = strText2 & "If Err <> 0 Then" & VbCrLf
        strText2 = strtext2 & "    Wscript.Echo " & chr(34) &  FilterValue & ":  Value not currently configured." & chr(34) & VbCrLf
        strText2 = strText2 & "Else" & VbCrLf       

        If strValueType = "REG_DWORD" Then
            strText = strText & "dwValue = " & objRecordSet.Fields.Item("DefaultValue") & VbCrLf
            strText = strText & "objReg.SetDWORDValue HKEY_CURRENT_USER, strKeyPath, ValueName, dwValue" & VbCrLf

            strText2 = strText2 & "    objReg.GetDWORDValue HKEY_CURRENT_USER, strKeyPath, ValueName, dwValue" & VbCrLf
            strText2 = strText2 & "    Wscript.Echo " & chr(34) & FilterValue & ": " & chr(34) & ", dwValue" & vbCrLf

        Else
            strText = strText & "strValue = " & chr(34) & objRecordSet.Fields.Item("DefaultValue") & chr(34) & VbCrLf
            strText = strText & "objReg.SetStringValue HKEY_CURRENT_USER, strKeyPath, ValueName, strValue" & VbCrLf

            strText2 = strText2 & "    objReg.GetStringValue HKEY_CURRENT_USER, strKeyPath, ValueName, strValue" & VbCrLf
            strText2 = strtext2 & "    Wscript.Echo " & chr(34) & FilterValue & ": " & chr(34) & ", strValue" & vbCrLf

        End If

        strText2 = strtext2 & "End If" 

        objRecordSet.MoveNext
    Loop
    HelpArea.InnerHTML = strHTML
    oFormTweakomatic.ScriptArea.innerText = strText
    oFormTweakomatic.RetrievalArea.innerText = strText2
    oFormTweakomatic.run_button.disabled = False
    oFormTweakomatic.run_button2.disabled = False
    oFormTweakomatic.save_button.disabled = False
    oFormTweakomatic.save_button2.disabled = False
    oFormTweakomatic.change_button.disabled = False
    oFormTweakomatic.Master_button.disabled = False
    oFormTweakomatic.Master_button2.disabled = False
End Sub


Sub RunConfigurationScript()
   Set objFS = CreateObject("Scripting.FileSystemObject")
   strTmpName = "temp_script.sm"
   Set objScript = objFS.CreateTextFile(strTmpName)
   objScript.Write oFormTweakomatic.ScriptArea.innerText
   objScript.Close
   Set objShell = CreateObject("WScript.Shell")
   strCmdLine = "wscript.exe //E:VBScript " & strTmpName
   objShell.Run strCmdLine
   strAction = "Configured value for " & oFormTweakomatic.TaskList.Value 
   ActionArea.InnerHTML = strAction
End Sub


Sub RunRetrievalScript()
   Set objFS = CreateObject("Scripting.FileSystemObject")
   strTmpName = "temp_script.sm"
   Set objScript = objFS.CreateTextFile(strTmpName)
   objScript.Write oFormTweakomatic.RetrievalArea.innerText
   objScript.Close
   Set objShell = CreateObject("WScript.Shell")
   strCmdLine = "wscript.exe //E:VBScript " & strTmpName
   objShell.Run strCmdLine
   strAction = "Retrieved value for " & oFormTweakomatic.TaskList.Value 
   ActionArea.InnerHTML = strAction
End Sub


Sub SaveConfigurationScript()
   Set objFSO = CreateObject("Scripting.FileSystemObject")
   strSaveFileName = InputBox("Please enter the complete path where you want to save your script (for example, C:\Scripts\MyScript.vbs).")
   If strSaveFileName = "" Then
      Exit Sub
   End If

   Set objFile = objFSO.CreateTextFile(strSaveFileName)
   objFile.WriteLine oFormTweakomatic.ScriptArea.innerText
   objFile.Close
   strAction = "Saved " & oFormTweakomatic.TaskList.Value & " to " & strSaveFileName
   ActionArea.InnerHTML = strAction
End Sub


Sub SaveRetrievalScript()
   Set objFSO = CreateObject("Scripting.FileSystemObject")
   strSaveFileName = InputBox("Please enter the complete path where you want to save your script (for example, C:\Scripts\MyScript.vbs).")
   If strSaveFileName = "" Then
      Exit Sub
   End If
   Set objFile = objFSO.CreateTextFile(strSaveFileName)
   objFile.WriteLine oFormTweakomatic.RetrievalArea.innerText
   objFile.Close
   strAction = "Saved " & oFormTweakomatic.TaskList.Value & " to " & strSaveFileName
   ActionArea.InnerHTML = strAction
End Sub


Sub ChangeValue()
    strCurrent = oFormTweakomatic.ScriptArea.innerText
    NewValue = InputBox("Please enter the new value: ")
    If NewValue = "" Then
        Exit Sub
    End If
    ScriptType = Split(strCurrent, vbCrLf, -1)
    If Left(ScriptType(6),2) = "dw" Then
        If Not IsNumeric(NewValue) Then
            Msgbox "You must enter a number when configuring DWORD registry values."
            Exit Sub
        End If
        ScriptType(6) = "dwValue = " & NewValue
    Else
        ScriptType(6) = "strValue = " & chr(34) & NewValue & chr(34)
    End If
    strReplace = Join(ScriptType, vbcrlf)
    oFormTweakomatic.ScriptArea.innerText = strReplace
    strAction = "Chnaged value for " & oFormTweakomatic.TaskList.Value & " to " & NewValue
    ActionArea.InnerHTML = strAction
End Sub


Sub ChangeMasterFile()
    If MasterFile = "" Then
        strCurrentFile = "Currently you do not have an master script. "
    Else
        strCurrentFile = "Your current master script is " & MasterFile & ". "
    End If
    strMessage = strCurrentFile & "Please enter the path to the new master script: "
    NewValue = InputBox(strMessage)
    If NewValue = "" Then
        Exit Sub
    End If
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    If objFSO.FileExists(NewValue) Then
        MasterFile = NewValue
        oFormTweakomatic.show_button.disabled = False
    Else
        CreateFile = Msgbox("This file does not exist. Would you like to create it",4)
        If CreateFile = vbYes Then
            objFSO.CreateTextFile(NewValue)
            MasterFile = NewValue
            oFormTweakomatic.show_button.disabled = False
        End If
    End If
    strAction = "Changed name of master script to " & NewValue
    ActionArea.InnerHTML = strAction
End Sub


Sub ChangeRetrievalFile()
    If RetrievalFile = "" Then
        strCurrentFile = "Currently you do not have a retrieval master script. "
    Else
        strCurrentFile = "You current retrieval master script is " & RetrievalFile & ". "
    End If
    strMessage = strCurrentFile & "Please enter the path to the new retrieval master script: "
    NewValue = InputBox(strMessage)
    If NewValue = "" Then
        Exit Sub
    End If
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    If objFSO.FileExists(NewValue) Then
        RetrievalFile  = NewValue
        oFormTweakomatic.show_button2.disabled = False
    Else
        CreateFile = Msgbox("This file does not exist. Would you like to create it",4)
        If CreateFile = vbYes Then
            objFSO.CreateTextFile(NewValue)
            RetrievalFile  = NewValue
            oFormTweakomatic.show_button2.disabled = False
        End If
    End If
    strAction = "Changed name of retrieval master script to " & NewValue
    ActionArea.InnerHTML = strAction
End Sub



Sub SetComputerName()
    strMessage = "Curently your scripts are using " & DefaultComputer & " as the default computer name. Please enter the new computer name. To run the script against the local computer, simply type a period (.) for the computer name: "
    NewValue = InputBox(strMessage)
    If NewValue = "" Then
        Exit Sub
    End If
    DefaultComputer = NewValue
    strAction = "Changed default computer name to " & NewValue
    ActionArea.InnerHTML = strAction
End Sub


Sub MasterConfigurationScript()
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    If MasterFile = "" Then
        strCurrentFile = "Currently you do not have a master script. "
        strMessage = strCurrentFile & "Please enter the path to the new master script: "
        NewValue = InputBox(strMessage)
        If NewValue = "" Then
            Exit Sub
        End If
            If objFSO.FileExists(NewValue) Then
            MasterFile  = NewValue
            oFormTweakomatic.show_button.disabled = False
        Else
            CreateFile = Msgbox("This file does not exist. Would you like to create it",4)
            If CreateFile = vbYes Then
                objFSO.CreateTextFile(NewValue)
                MasterFile  = NewValue
                oFormTweakomatic.show_button.disabled = False
            Else
                Exit Sub
            End If
        End If
    End If
    Set objFile = objFSO.OpenTextFile(MasterFile, 8)
    objFile.WriteLine Chr(39) & "   " & oFormTweakomatic.TaskList.Value
    objFile.WriteLine oFormTweakomatic.ScriptArea.innerText
    objFile.WriteLine vbCrLf & vbCrLf
    objFile.Close
    strAction = "Appended " & oFormTweakomatic.TaskList.Value & " to " & MasterFile
    ActionArea.InnerHTML = strAction
End Sub


Sub MasterRetrievalScript()
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    If RetrievalFile = "" Then
        strCurrentFile = "Currently you do not have a retrieval master script. "
        strMessage = strCurrentFile & "Please enter the path to the new master script: "
        NewValue = InputBox(strMessage)
        If NewValue = "" Then
            Exit Sub
        End If
            If objFSO.FileExists(NewValue) Then
            RetrievalFile  = NewValue
            oFormTweakomatic.show_button2.disabled = False
        Else
            CreateFile = Msgbox("This file does not exist. Would you like to create it",4)
            If CreateFile = vbYes Then
                objFSO.CreateTextFile(NewValue)
                RetrievalFile  = NewValue
                oFormTweakomatic.show_button2.disabled = False
            Else
                Exit Sub
            End If
        End If
    End If
    Set objFile = objFSO.OpenTextFile(RetrievalFile, 8)
    objFile.WriteLine Chr(39) & "   " & oFormTweakomatic.TaskList.Value
    objFile.WriteLine oFormTweakomatic.RetrievalArea.innerText
    objFile.WriteLine vbCrLf & vbCrLf
    objFile.Close
    strAction = "Appended " & oFormTweakomatic.TaskList.Value & " to " & RetrievalFile
    ActionArea.InnerHTML = strAction
End Sub


Sub ShowConfigurationScript()
   Set objShell = CreateObject("WScript.Shell")
   strCmdLine = "notepad.exe " & MasterFile
   objShell.Run strCmdLine
   strAction = "Opened file " & MasterFile & " in Notepad"
   ActionArea.InnerHTML = strAction
End Sub



Sub ShowRetrievalScript()
   Set objShell = CreateObject("WScript.Shell")
   strCmdLine = "notepad.exe " & RetrievalFile
   objShell.Run strCmdLine
   strAction = "Opened file " & RetrievalFile & " in Notepad"
   ActionArea.InnerHTML = strAction
End Sub