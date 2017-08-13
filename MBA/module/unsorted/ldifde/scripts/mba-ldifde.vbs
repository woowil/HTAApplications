

'Option Explicit
	'Dim enumAttributes, objclassinputA, LaunchCmd, GenerateSyntax, RunLdifde, strSyntaxA
	Sub Window_Onload_LDIFDE
		Dim objDialogWindow, DataList, DataListA, objDSE, _
				objSchema, objClass, objReturn, strHTML, strClass, strCategory, objNetwork, objNetUser, _
				objNetDomain, strHTMLSSPI, y, z
		
		'- Constants used in recordset fields
		Const adVarChar = 200 'set datatype to variant
		Const MaxCharacters = 255 'set maximum characters to 255
		
		'- Set initial zero value for dictionary keys.
		y=0
		z=0
		
		'On Error Resume Next
		'- Create the ADO Recordset instance and add two fields to the recordset:
		'- objClass and objCategory
		Set DataList = CreateObject("ADOR.Recordset")
		DataList.Fields.Append "objClass", adVarChar, MaxCharacters
		DataList.Open
		Set DataListA = CreateObject("ADOR.Recordset")
		DataListA.Fields.Append "objCategory", adVarChar, MaxCharacters
		DataListA.Open
		'- Declare array and List Schema Objects. All object headers are listed for
		'- informational value although only two are used (listed below).
		Dim arrSchemaObj(14)
		arrSchemaObj(0) = "lDAPDisplayName"
		arrSchemaObj(1) = "cn"
		arrSchemaObj(2) = "adminDisplayName"
		arrSchemaObj(3) = "distinguishedName"
		arrSchemaObj(4) = "adminDescription"
		arrSchemaObj(5) = "governsID"
		arrSchemaObj(6) = "rDNAttID"
		arrSchemaObj(7) = "objectClassCategory"
		arrSchemaObj(8) = "subClassOf"
		arrSchemaObj(9) = "defaultObjectCategory"
		arrSchemaObj(10) = "defaultHidingValue"
		arrSchemaObj(11) = "defaultSecurityDescriptor"
		arrSchemaObj(12) = "systemOnly"
		arrSchemaObj(13) = "systemFlags"
		arrSchemaObj(14) = "objectCategory"
		'- Bind to RootDSE and build an ADsPath to the Schema container.
		Set objDSE = GetObject("LDAP://rootDSE")
		Set objSchema = GetObject("LDAP://" & objDSE.Get("schemaNamingContext"))
		
		On Error Resume Next
		'- Enum class objects into array.
		For Each objChild In objSchema
			If objChild.Class = "classSchema" Then
				For i = LBound(arrSchemaObj) To UBound(arrSchemaObj)
					objReturn = objChild.Get(arrSchemaObj(i))
					If Err <> 0 Then
						Err.Clear
						objReturn = ""
					End If
					'- Add lDAPDisplayName schema objects to the recordset.
					If i=0 Then
						DataCompare.Add y, objReturn
						y=y+1
						DataList.AddNew
						DataList("objClass") = objReturn
						DataList.Update
					End If
					'- Add defaultObjectCategory schema objects to the recordset.
					If i=9 Then
						DataCompare.Add z, objReturn
						z=z+1
						DataListA.AddNew
						DataListA("objCategory") = objReturn
						DataListA.Update
					End If
				Next
			End If
		Next
		
		'- Sort the recordset to complete schema class select dropdown.
		DataList.Sort = "objClass"
		DataList.MoveFirst
		'- Enumerate associated attributes for selected class upon mouse click.
		strHTML = "<select name=""objclassinput"" size=""4"" onClick=""enumAttributes objclassinput"">"
		Do Until DataList.EOF
			strClass = DataList.Fields.Item("objClass")
			If strClass = "user" Then
				strHTML=strHTML & "<option selected value=" & Chr(34) & strClass & Chr(34) & ">" & _
				strClass & "</option>" & VbCrLf
			Else
				strHTML=strHTML & "<option value=" & Chr(34) & strClass & Chr(34) & ">" & _
				strClass & "</option>" & VbCrLf
			End If
			DataList.MoveNext
		Loop
		strHTML = strHTML & "</select>"
		objClassList.InnerHTML = strHTML
		DataListA.Sort = "objCategory"
		DataListA.MoveFirst
		'- Sort the recordset to complete schema category select dropdown.
		strHTML = "<select style=""width:350px"" name=""objcategoryinput"" size=""4"">"
		Do Until DataListA.EOF
			strCategory = DataListA.Fields.Item("objCategory")
			strHTML=strHTML & "<option value=" & Chr(34) & strCategory & Chr(34) & ">" & _
			strCategory & "</option>" & VBCrLf
			DataListA.MoveNext
		Loop
		strHTML = strHTML & "</select>"
		objCategoryList.InnerHTML = strHTML
		'- Enumerate current user's distinguishedName attribute to complete authentication options.
		Set objSysInfo = CreateObject("ADSystemInfo")
		strHTMLDN = "<input name=""userdn"" type=""text"" value=" & Chr(34) & objSysInfo.UserName & _
		Chr(34) & " size=""40"">"
		objUserDN.InnerHTML = strHTMLDN		
		'- Enumerate current user and NetBIOS domain information to complete authentication options.
		Set objNetwork = CreateObject("WScript.Network")
		objNetUser = objNetwork.UserName
		objNetDomain = objNetwork.UserDomain
		strHTMLSSPI = "<input name=""username"" type=""text"" value=" & objNetUser & " size=""15"">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
		strHTMLSSPI = strHTMLSSPI & "<input name=""domain"" type=""text"" value=" & objNetDomain & " size=""15"">"
		objNetDom.InnerHTML = strHTMLSSPI
	End Sub
	
	Sub enumAttributes(objclassinputA)
		Dim objClass, strHTMLatt, strHTMLattincl, strHTMLattomit, strPropName
		
		'- Enum attributes depending upon what schema class is currectly selected.
		Set objClass = GetObject("LDAP://schema/" & objclassinputA.Value)
		strHTMLatt = "<select name=""attributeinput"" size=""1"">"
		strHTMLattincl = "<select name=""attributeinclude"" size=""4"" multiple>"
		strHTMLattomit = "<select name=""attributeomit"" size=""4"" multiple>"
		For Each strPropName In objClass.MandatoryProperties
			strHTMLatt = strHTMLatt & "<option value=" & Chr(34) & strPropName & Chr(34) & ">" & _
			strPropName & "</option>" & VbCrLf
			strHTMLattincl = strHTMLattincl & "<option value=" & Chr(34) & strPropName & Chr(34) & ">" & _
			strPropName & "</option>" & VbCrLf
			strHTMLattomit = strHTMLattomit & "<option value=" & Chr(34) & strPropName & Chr(34) & ">" & _
			strPropName & "</option>" & VbCrLf
		Next
		For Each strPropName In objClass.OptionalProperties
			strHTMLatt = strHTMLatt & "<option value=" & Chr(34) & strPropName & Chr(34) & ">" & _
			strPropName & "</option>" & VbCrLf
			strHTMLattincl = strHTMLattincl & "<option value=" & Chr(34) & strPropName & Chr(34) & ">" & _
			strPropName & "</option>" & VbCrLf
			strHTMLattomit = strHTMLattomit & "<option value=" & Chr(34) & strPropName & Chr(34) & ">" & _
			strPropName & "</option>" & VbCrLf
		Next
		strHTMLatt = strHTMLatt & "</select>"
		strHTMLattincl = strHTMLattincl & "</select>"
		strHTMLattomit = strHTMLattomit & "</select>"
		objAttributeList.InnerHTML = strHTMLatt
		listincludeinput.InnerHTML = strHTMLattincl
		listomitinput.InnerHTML = strHTMLattomit
	End Sub
	
	Sub LaunchCmd '- Launch the ldifde command with selections.
		Dim objName, strFilename, strServername, strreplacedn, strlogfile, strport, strtimeout, _
				strrootdn, strfilter, intfilter, strobjectclass, strobjectcategory, strattribute, _
				strFilterVal, objRadio, objRadioButton, strsearchscope, attributeincludeval, strlistinclude, _
				attributeomitval, strlistomit, strthread, strsimple, strsspi, strSyntax
				
		Set objName = document.oFormLDIFDE
		If objName.importmode.Checked Then strImport = objName.importmode.Value
		If objName.filename.Checked Then
			strFilename = objName.filename.Value & " " & """" & objName.path.Value & """"
		End If
		If objName.servername.Checked Then 
			strServername = objName.servername.Value & " " & objName.servernameinput.Value
		End If
		If objName.replacedn.Checked Then
			strreplacedn = objName.replacedn.Value & " " & """" & objName.replacedninput.Value & """"
		End If
		If objName.verbose.Checked Then strverbose = objName.verbose.Value
		If objName.logfile.Checked Then
			strlogfile = objName.logfile.Value & " " & """" & objName.logfileinput.Value & """"
		End If
		If objName.port.Checked Then
			strport = objName.port.Value & " " & objName.portinput.Value
		End If
		If objName.unicode.Checked Then strunicode = objName.unicode.Value
		If objName.timeout.Checked Then
			strtimeout = objName.timeout.Value & " " & objName.timeoutinput.Value
		End If
		If objName.sasl.Checked Then strsasl = objName.sasl.Value
		If objName.rootdn.Checked Then
			strrootdn = objName.rootdn.Value & " " & """" & objName.rootdninput.Value & """"
		End If
		intFilter = 0
		If objName.filter.Checked Then
			strfilter = " " & objName.filter.Value
			intFilter = intFilter+1
		End If
		If objName.objectclass.Checked Then
			objName.filter.Checked = True
			strobjectclass = "(objectClass=" & objName.objclassinput.Value & ")"
			intFilter = intFilter+2
		End If
		If objName.objectcategory.Checked Then
			objName.filter.Checked = True
			strobjectcategory = "(objectCategory=" & objName.objcategoryinput.Value & ")"
			intFilter = intFilter+2
		End If
		If objName.attribute.Checked Then
			objName.filter.Checked = True
			strattribute = "(" & objName.attributeinput.Value & "=" & objName.attributename.Value & ")"
			intFilter = intFilter+2
		End If
		If intFilter > 2 Then
			strFilterVal = strfilter & " " & Chr(34) & " (&" & strobjectclass & strobjectcategory & _
			strattribute & ")" & Chr(34)
		End If
		If intFilter = 1 Then
			strFilterVal =  strfilter & " " & Chr(34) & strobjectclass & strobjectcategory & strattribute & Chr(34)
		End If
		If objName.searchscopechk.Checked Then
		Set objRadio = objName.Elements("searchscope")
		For Each objRadioButton In objRadio
			If objRadioButton.Checked = True Then
				strsearchscope = objName.searchscopechk.Value & " " & objRadioButton.Value
				Exit For
			End If
		Next
		End If
		If objName.listinclude.Checked Then 
		For i = 0 to (objName.attributeinclude.Options.Length - 1)
        If (objName.attributeinclude.Options(i).Selected) Then
            attributeincludeval = attributeincludeval & objName.attributeinclude.Options(i).Value & ","
        End If
    Next
			strlistinclude = objName.listinclude.Value & " " & """" & attributeincludeval & """"
		End If
		If objName.listomit.Checked Then 
		For i = 0 to (objName.attributeomit.Options.Length - 1)
        If (objName.attributeomit.Options(i).Selected) Then
            attributeomitval = attributeomitval & objName.attributeomit.Options(i).Value & ","
        End If
    Next
			strlistomit = objName.listomit.Value & " " & """" & attributeomitval & """"
		End If
		If objName.paged.Checked Then strpaged = objName.paged.Value
		If objName.samlogic.Checked Then strsamlogic = objName.samlogic.Value
		If objName.binaryexport.Checked Then strbinaryexport = objName.binaryexport.Value
		If objName.ignoreerrors.Checked Then strignoreerrors = objName.ignoreerrors.Value
		If objName.lazycommit.Checked Then strlazycommit = objName.lazycommit.Value
		If objName.thread.Checked Then 
			strthread = objName.thread.Value & " " & objName.threadinput.Value
		End If
		If objName.simple.Checked Then 
			strsimple = " " & """" & objName.userdn.Value & """" & " " & objName.userdnpassword.Value
		End If
		If objName.sspi.Checked Then 
			strsspi = " " & """" & objName.username.Value & """" & " " & objName.domain.Value & _
			" " & objName.usernamepass.Value
		End If
		
		strSyntax = strImport & strFilename & strServername & strreplacedn & strverbose & strlogfile & _
								strport & strunicode & strtimeout & strsasl & strrootdn & strFilterVal & strsearchscope & strlistinclude & _
								strlistomit & strpaged & strsamlogic & strbinaryexport & strignoreerrors & strlazycommit & _
								strtrhread & strsimple & strsspi
		RunLdifde strSyntax
	End Sub

	Sub GenerateSyntax '- Generate ldifde syntax and copy to clipboard.
		Dim objName, strFilename, strServername, strreplacedn, strlogfile, strport, strtimeout, _
				strrootdn, strfilter, intfilter, strobjectclass, strobjectcategory, strattribute, _
				strFilterVal, objRadio, objRadioButton, strsearchscope, attributeincludeval, strlistinclude, _
				attributeomitval, strlistomit, strthread, strsimple, strsspi, strSyntax
				
		Set objName = document.oFormLDIFDE
		If objName.importmode.Checked Then strImport = objName.importmode.Value
		If objName.filename.Checked Then
			strFilename = objName.filename.Value & " " & """" & objName.path.Value & """"
		End If
		If objName.servername.Checked Then 
			strServername = objName.servername.Value & " " & objName.servernameinput.Value
		End If
		If objName.replacedn.Checked Then
			strreplacedn = objName.replacedn.Value & " " & """" & objName.replacedninput.Value & """"
		End If
		If objName.verbose.Checked Then strverbose = objName.verbose.Value
		If objName.logfile.Checked Then
			strlogfile = objName.logfile.Value & " " & """" & objName.logfileinput.Value & """"
		End If
		If objName.port.Checked Then
			strport = objName.port.Value & " " & objName.portinput.Value
		End If
		If objName.unicode.Checked Then strunicode = objName.unicode.Value
		If objName.timeout.Checked Then
			strtimeout = objName.timeout.Value & " " & objName.timeoutinput.Value
		End If
		If objName.sasl.Checked Then strsasl = objName.sasl.Value
		If objName.rootdn.Checked Then
			strrootdn = objName.rootdn.Value & " " & """" & objName.rootdninput.Value & """"
		End If
		intFilter = 0
		If objName.filter.Checked Then
			strfilter = " " & objName.filter.Value
			intFilter = intFilter+1
		End If
		If objName.objectclass.Checked Then
			objName.filter.Checked = True
			strobjectclass = "(objectClass=" & objName.objclassinput.Value & ")"
			intFilter = intFilter+2
		End If
		If objName.objectcategory.Checked Then
			objName.filter.Checked = True
			strobjectcategory = "(objectCategory=" & objName.objcategoryinput.Value & ")"
			intFilter = intFilter+2
		End If
		If objName.attribute.Checked Then
			objName.filter.Checked = True
			strattribute = "(" & objName.attributeinput.Value & "=" & objName.attributename.Value & ")"
			intFilter = intFilter+2
		End If
		If intFilter > 2 Then
			strFilterVal = strfilter & " " & Chr(34) & " (&" & strobjectclass & strobjectcategory & _
			strattribute & ")" & Chr(34)
		End If
		If intFilter = 1 Then
			strFilterVal =  strfilter & " " & Chr(34) & strobjectclass & strobjectcategory & strattribute & Chr(34)
		End If
		If objName.searchscopechk.Checked Then
		Set objRadio = objName.Elements("searchscope")
		For Each objRadioButton In objRadio
			If objRadioButton.Checked = True Then
				strsearchscope = objName.searchscopechk.Value & " " & objRadioButton.Value
				Exit For
			End If
		Next
		End If
		If objName.listinclude.Checked Then 
		For i = 0 to (objName.attributeinclude.Options.Length - 1)
        If (objName.attributeinclude.Options(i).Selected) Then
            attributeincludeval = attributeincludeval & objName.attributeinclude.Options(i).Value & ","
        End If
    Next
			strlistinclude = objName.listinclude.Value & " " & """" & attributeincludeval & """"
		End If
		If objName.listomit.Checked Then 
		For i = 0 to (objName.attributeomit.Options.Length - 1)
        If (objName.attributeomit.Options(i).Selected) Then
            attributeomitval = attributeomitval & objName.attributeomit.Options(i).Value & ","
        End If
    Next
			strlistomit = objName.listomit.Value & " " & """" & attributeomitval & """"
		End If
		If objName.paged.Checked Then strpaged = objName.paged.Value
		If objName.samlogic.Checked Then strsamlogic = objName.samlogic.Value
		If objName.binaryexport.Checked Then strbinaryexport = objName.binaryexport.Value
		If objName.ignoreerrors.Checked Then strignoreerrors = objName.ignoreerrors.Value
		If objName.lazycommit.Checked Then strlazycommit = objName.lazycommit.Value
		If objName.thread.Checked Then 
			strthread = objName.thread.Value & " " & objName.threadinput.Value
		End If
		If objName.simple.Checked Then 
			strsimple = " " & """" & objName.userdn.Value & """" & " " & objName.userdnpassword.Value
		End If
		If objName.sspi.Checked Then 
			strsspi = " " & """" & objName.username.Value & """" & " " & objName.domain.Value & _
			" " & objName.usernamepass.Value
		End If
		
		strSyntax = strImport & strFilename & strServername & strreplacedn & strverbose & strlogfile & _
								strport & strunicode & strtimeout & strsasl & strrootdn & strFilterVal & strsearchscope & strlistinclude & _
								strlistomit & strpaged & strsamlogic & strbinaryexport & strignoreerrors & strlazycommit & _
								strtrhread & strsimple & strsspi
		objName.CopytoClip.Value = "ldifde " & strSyntax
		strCopy = objName.CopytoClip.Value
    document.parentwindow.clipboardData.SetData "text", strCopy
    strHTMLCopied = "<font style=""font-family: Tahoma; color: red; font-weight: bold; font-size: 12px"">"
    strHTMLCopied = strHTMLCopied & "The generated code has been copied to the clipboard</font>"
    copied.InnerHTML = strHTMLCopied
	End Sub
	
	Sub RunLdifde(strSyntaxA) 'Run ldifde in command window with currently selected options.
		Dim strRunCmd, strMsgBox, objShell
		
		strRunCmd = "ldifde " & strSyntaxA
		strMsgBox = MsgBox("Do you want to run the following command? Selecting Cancel will exit" & _
											" the application." & VbCrLf & VbCrLf & strRunCmd, 35, "Run LDIFDE?")
		If strMsgBox = 7 Then
			MsgBox "Action Cancelled..." & VbCrLf & "Refreshing window."
			document.location.reload(True)
		ElseIf strMsgBox = 2 Then
			window.close
		Else
			Set objShell = CreateObject("WScript.Shell")
			objShell.Run "Cmd.exe /k " & strRunCmd, 3, False
		End If
	End Sub
	