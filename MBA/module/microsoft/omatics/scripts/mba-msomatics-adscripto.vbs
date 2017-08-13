'***************************************************************************
'* This routine runs when the application loads. The interface is sized
'* and UI elements appear.
'*************************************************************************** 

Sub Window_Onload_ADScripto()
  strHTML = "<select onChange=""TaskSelectPulldownCheck()"" name=TaskSelectPulldown style=""width:160px;"">" & _
              "<option value=""PulldownMessage"">Select a task"
  arrTasks = Array("Create an Object","Write an Object","Read an Object","Delete an Object")

  For Each obj in arrTasks   
    strHTML = strHTML & "<option value= " & chr(34) & _
      obj & chr(34) & ">" & obj
  Next
  strHTML = strHTML & "</select>"
  
  'Create a list of the ADSI classes available for script creation. 
  'A small number of Class names in the collection will populate the
  'classes pulldown.
  strHTML = strHTML & " &nbsp; <select onChange=""TaskSelectPulldownCheck()"" name=ClassesPulldown style=""width:160px;"">" & _
              "<option value=""PulldownMessage"">Select a class"
  arrClassObjects = Array("user","computer","contact","group","organizationalUnit")
  For Each obj in arrClassObjects   
    strHTML = strHTML & "<option value= " & chr(34) & _
      obj & chr(34) & ">" & obj
  Next
  strHTML = strHTML & "</select>"
  
  ooDivADScriptADSI.innerHTML = strHTML
  strImptNote = "EzAD Scriptomatic is an ADSI Scripting learning tool. " & _
               "The tool creates example scripts that read, write and " & _
               "modify Active Directory data. To successfully run " & _
               "scripts created with EzAD Scriptomatic, you must:<br>" & _
               "<li>Have Administrator access to Active Directory, and " & _
               "<li>Be logged on to the target Active Directory domain.<br><br>" & _
               "You should not run scripts created with EzAD Scriptomatic" & _
               "against a production domain without first testing the"& _
               "scripts in your designated testing environment."
  Call InitialUIState
  oDivADScriptCode.innerHTML = strImptNote  
 End Sub 

'***************************************************************************
'* These subroutines control the state of the user interface. Each 
'* routine includes descriptive text.
'*************************************************************************** 
'This is the state the HTA UI elements should be in 
'before anything is selected.
Sub InitialUIState()
   oFormADScripto.TaskSelectPulldown.disabled = False
   oFormAction.action(0).disabled = True
   oFormAction.action(1).disabled = True
   oFormADScripto.runNotes_button.value = ""
   oFormADScripto.classesPulldown.selectedIndex = 0
   oFormADScripto.classesPulldown.disabled = True
   oDivADScriptCode.innerHTML = ""
   oFormADScripto.notes_button.value = ""
End Sub

'If the operator selects a different task, reset the UI in 
'preperation for selecting a class. 
Sub ResetForClassesPullDown()
  oFormADScripto.ClassesPullDown.Disabled = False
  oDivADScriptCode.InnerHTML= ""
  oFormADScripto.notes_button.value = ""
  oFormADScripto.runNotes_button.value = ""
End Sub

'If the TaskSelectPulldown is not set to Select a task
'then enable the ClassesPulldown. Otherwise, disable
'the ClassesPulldown. 
Sub TaskSelectPulldownCheck
  Select Case oFormADScripto.TaskSelectPulldown.SelectedIndex
    Case "0"
      Call InitialUIState
    Case "1"
      Call ResetForClassesPullDown
      Call CreateCreateScript
    Case "2"
      Call ResetForClassesPullDown
      Call CreateWriteScript
    Case "3"
      Call ResetForClassesPullDown
      Call CreateReadScript
    Case "4"
      Call ResetForClassesPullDown
      Call CreateDeleteScript
  End Select
End Sub

'Once a script is generated, enable the Running This Script, Run, and Save buttons. 
Sub FinalUIState
   oFormADScripto.runNotes_button.value="Click here to read this before running the " & _
    lcase(oFormADScripto.TaskSelectPulldown.Value) & " - " & oFormADScripto.ClassesPulldown.Value & " script"
   oFormAction.action(0).disabled = False
   oFormAction.action(1).disabled = False
End Sub

Sub ChangePointer
  'If window.event.srcElement.id = "notesbuttonrun" Then
   ' oFormADScripto.runNotes_button.style.cursor="hand"
 ' ElseIf window.event.srcElement.id = "notesbutton" Then
  '  oFormADScripto.notes_button.style.cursor="hand"
  'End If
End Sub

'***************************************************************************
'* These functions are called by other functions that generate the code for  
'* the Read an Object task. Each function in this section generates code
'* based on attribute definitions.
'***************************************************************************
Function StrReadCodeSV(strPageName,arrName)
  strHTML = strHTML & "WScript.Echo VbCrLf &amp; " & chr(34) & "** " & strPageName & " Properties Page**" & chr(34) & chr(10)
  strHTML = strHTML & "WScript.Echo " & chr(34) & "** (Single-Valued Attributes) **" & chr(34) & chr(10)
  For each attrib in arrName
    'Added the Instr function here because the COM+ attribute, msCOM-UserPartitionSetLink
    'contains a hyphen, which a variable name should not contain.
    If Instr(1,attrib,"-") = 0 Then
      strHTML = strHTML & _
        "str" & attrib & " = objItem.Get(" & chr(34) & attrib & chr(34) & ")" & chr(10)
      strHTML = strHTML & _
        "WScript.Echo " & chr(34) & attrib & ": " & chr(34) & " &amp; str" & attrib & chr(10)
    Else
      strHTML = strHTML & _
        "WScript.Echo " & chr(34) & attrib & ": " & chr(34) & chr(10) & _
        "WScript.Echo " & chr(34) & " " & chr(34) & " &amp; objItem.Get(" & chr(34) & attrib & chr(34) & ")" & chr(10)
    End If
  Next
  strReadCodeSV = strHTML & chr(10)
End Function

Function StrReadCodeMV(strPageName,arrName)
  strHTML = strHTML & "WScript.Echo VbCrLf &amp; " & chr(34) & "** " & strPageName & " Properties Page**" & chr(34) & chr(10)
  strHTML = strHTML & "WScript.Echo " & chr(34) & "** (MultiValued Attributes) **" & chr(34) & chr(10)
  For Each attrib in arrName
    strHTML = strHTML & _
      "str" & attrib & " = objItem.GetEx(" & chr(34) & attrib & chr(34) & ")" & chr(10)
    strHTML = strHTML & _
      "WScript.Echo " & chr(34) & attrib & ":" & chr(34) & chr(10)
    strHTML = strHTML & _
      "For Each Item in str" & attrib & chr(10)
    strHTML = strHTML & _
      " WScript.Echo vbTab &amp; Item" & chr(10)
    strHTML = strHTML & "Next" & chr(10)
   Next 
   StrReadCodeMV = strHTML & chr(10)
End Function

'For reading attributes stored as integers containing bit flags
Function IntReadCode(strPageName,attrib,arrConstant,arrValue)
  strHTML = strHTML & "WScript.Echo VbCrLf &amp; " & chr(34) & "** " & strPageName & " Properties Page**" & chr(34) & chr(10)
  strHTML = strHTML & "WScript.Echo " & chr(34) & "** (The " & attrib & " attribute) **" & chr(34) & chr(10)
  strHTML = strHTML & "Set objHash = CreateObject(" & chr(34) & "Scripting.Dictionary" & chr(34) & ")" & chr(10)
  i = 0
  For Each constant in arrConstant
    strHTML = strHTML & "objHash.Add " & chr(34) & constant & chr(34) & ", " & arrValue(i) & chr(10) 
    i = 1 + i
  Next
  strHTML = strHTML & _
    "int" & attrib & " = objItem.Get(" & chr(34) & attrib & chr(34) & ")" & chr(10)
  strHTML = strHTML & _
    "For Each Key in objHash.Keys" & chr(10)
  strHTML = strHTML & _ 
    "  If objHash(Key) And " & "int" & attrib & " Then" & chr(10)
  strHTML = strHTML & _
    "    WScript.Echo Key &amp; " & chr(34) & " is enabled." & chr(34) & chr(10)
  strHTML = strHTML & "  Else" & chr(10)
  strHTML = strHTML & _
    "    WScript.Echo Key &amp; " & chr(34) & " is disabled." & chr(34) & chr(10)
  strHTML = strHTML & "  End If" & chr(10)
  strHTML = strHTML & "Next"
  IntReadCode = strHTML & chr(10)
End Function

Function ReadPropertiesSimple(strPageName,interfaceName,arrProp)
  strHTML = strHTML & "WScript.Echo VbCrLf &amp; " & chr(34) & "** " & strPageName & " Properties Page**" & chr(34) & chr(10)
  strHTML = strHTML & "WScript.Echo " & chr(34) & "** (The " & interfaceName & " interface) **" & chr(34) & chr(10)
  For Each prop in arrProp
    strHTML = strHTML & "WScript.Echo " & _
      chr(34) & prop & ": " & chr(34) & " &amp; _" & chr(10) & _
      "  objItem." & prop & chr(10)
  Next
  ReadPropertiesSimple = strHTML & chr(10)
End Function

'***************************************************************************
'* These functions are called by other functions that generate the code for  
'* the Write an Object task. Each function in this section generates code
'* based on attribute definitions.
'***************************************************************************
Function StrWriteCodeSV(strPageName,arrName,strValue)
  If strValue = "VALUE" Then
    strHTML = strHTML & "WScript.Echo VbCrLf &amp; " & chr(34) & "** " & strPageName & " Properties Page**" & chr(34) & chr(10)
    strHTML = strHTML & "WScript.Echo " & chr(34) & "** (Writing Single-Valued Attributes) **" & chr(34) & chr(10)
  Else
    strHTML = strHTML & "WScript.Echo VbCrLf &amp; " & chr(34) & "**Non-standard value on " & strPageName & " Properties Page**" & chr(34) & chr(10)
    strHTML = strHTML & "'See Script Notes for information on setting this value." & chr(10)
  End If
  For each attrib in arrName
    strHTML = strHTML & _
      "objItem.Put " & chr(34) & attrib & chr(34) & ", " & _
        chr(34) & strValue & chr(34) & chr(10)
    strHTML = strHTML & "objItem.SetInfo" & chr(10)
  Next
  strWriteCodeSV = strHTML & chr(10)
End Function

Function StrWriteCodeMV(strPageName,arrName,strValue)
 If strValue = "VALUE" Then
  strHTML = strHTML & "WScript.Echo VbCrLf &amp; " & chr(34) & "** " & strPageName & " Properties Page**" & chr(34) & chr(10)
  strHTML = strHTML & "WScript.Echo " & chr(34) & "** (Writing MultiValued Attributes) **" & chr(34) & chr(10)
 Else
  strHTML = strHTML & "WScript.Echo VbCrLf &amp; " & chr(34) & "**Non-standard value on " & strPageName & " Properties Page**" & chr(34) & chr(10)
  strHTML = strHTML & "'See Script Notes for information on setting this value." & chr(10)
 End If
 For each attrib in arrName
  strHTML = strHTML & _
    "objItem.PutEx ADS_PROPERTY_UPDATE, " & chr(34) & attrib & chr(34) & ", _" & chr(10) & _
      "  Array(" & chr(34) & strValue & "1" & chr(34) & _
      ", " & chr(34) & strValue & "2" & chr(34) & _ 
      ", " & chr(34) & "..." & strValue & "n" & chr(34) & ")" & chr(10)
    strHTML = strHTML & "objItem.SetInfo" & chr(10)
  Next
  strWriteCodeMV = strHTML & chr(10)
End Function

'***************************************************************************
'* These functions write script code to the code window for the Read 
'* an Object task. Each function varies based on the selected class. 
'* The function name describes the class it supports.
'***************************************************************************
Function UserAttribsToRead
  'All attributes on the General Properites Page
  arrSVStringAttribsGP = Array("name", "givenName","initials", _
    "sn","displayName","description","physicalDeliveryOfficeName", _
    "telephoneNumber","mail","wWWHomePage")
  strHTML = strReadCodeSV("General",arrSVStringAttribsGP) 
  
  arrMVStringAttribsGP = Array("otherTelephone", "url")
  strHTML = strHTML & strReadCodeMV("General",arrMVStringAttribsGP)
  'End General Properties Page
  
  'All attributes on the Address Properties Page
  arrSVStringAttribsAP = Array("streetAddress", "l", "st", _
    "postalCode", "c")
  strHTML = strHTML & strReadCodeSV("Address",arrSVStringAttribsAP)

  arrMVStringAttribsAP = Array("postOfficeBox")
  strHTML = strHTML & strReadCodeMV("Address",arrMVStringAttribsAP)
  'End Address Properties Page

  'Selected attributes on the Account Properties Page
  arrSVStringAttribsAcP = Array("userPrincipalName", "dc", _
    "sAMAccountName", "userWorkstations") 
  strHTML  = strHTML & strReadCodeSV("Account",arrSVStringAttribsAcP)
  
  'Read the bit flags in userAccountControl
  arrUACConstants = Array("ADS_UF_SMARTCARD_REQUIRED", _
    "ADS_UF_TRUSTED_FOR_DELEGATION", "ADS_UF_NOT_DELEGATED", _
    "ADS_UF_USE_DES_KEY_ONLY","ADS_UF_DONT_REQUIRE_PREAUTH")
  arrUACValues = Array("&h40000", "&h80000", "&h100000", _
    "&h200000", "&h400000")
  strHTML = strHTML & IntReadCode("Account","userAccountControl", _
    arrUACConstants,arrUACValues)
  'End read the bit flags in userAccountControl
  
  'Read the IsAccountLocked property
  strHTML = strHTML & "If objItem.IsAccountLocked = True Then" & chr(10)
  strHTML = strHTML & _
    "  WScript.Echo " & chr(34) & "ADS_UF_LOCKOUT is enabled" & chr(34) & chr(10)
  strHTML = strHTML & "Else" & chr(10)
  strHTML = strHTML & _
    "  WScript.Echo " & chr(34) & "ADS_UF_LOCKOUT is disabled" & chr(34) & chr(10)
  strHTML = strHTML & "End If" & chr(10) & chr(10)
  'End read the IsAccountLocked property
  
  'Read the AccountExpirationDate property
    strHTML = strHTML & "If err.Number = -2147467259 OR _" & chr(10) & _
     "  objItem.AccountExpirationDate = " & chr(34) & "1/1/1970" & chr(34) & _
     " Then" & chr(10)
    strHTML = strHTML & "  WScript.Echo " & chr(34) & _
      "Account doesn't expire." & chr(34) & chr(10)
    strHTML = strHTML & "Else" & chr(10)
    strHTML = strHTML & "  WScript.Echo " & chr(34) & _
      "Account expires on: " & chr(34) & " &amp; " & _
      "objItem.AccountExpirationDate" & chr(10)
    strHTML = strHTML & "End If" & chr(10) & chr(10)
  'End read the AccountExpirationDate property
  'End Account Properties Page 
  
  'All attributes on the Profile Properties Page
  arrSVStringAttribsPrP = Array("profilePath", "scriptPath", _
    "homeDirectory", "homeDrive")
  strHTML = strHTML & strReadCodeSV("Profile",arrSVStringAttribsPrP)
  'End Profile Properties Page 
  
  'All attributes on the Telephones Properties Page
  arrSVStringAttribsTele = Array("homePhone","pager", _
    "mobile","facsimileTelephoneNumber","ipPhone", "info")
  strHTML = strHTML & strReadCodeSV("Telephone",arrSVStringAttribsTele)
  
  arrMVStringAttribsTele = Array("otherHomePhone","otherPager", _
    "otherMobile","otherFacsimileTelephoneNumber","otherIpPhone")
  strHTML = strHTML & strReadCodeMV("Telephone",arrMVStringAttribsTele)
  'End Telephones Properties Page 
  
  'All attributes on the Organization Properties Page
  arrSVStringAttribsOrg = Array("title","department", _
    "company","manager")
  strHTML = strHTML & strReadCodeSV("Organization",arrSVStringAttribsOrg)
  
  arrMVStringAttribsOrg = Array("directReports")
  strHTML = strHTML & strReadCodeMV("Organization",arrMVStringAttribsOrg)
  'End Organization Properties Page 
  
  'All settings on the Environment Properties Page
  arrProperties = Array("TerminalServicesInitialProgram", _
    "TerminalServicesWorkDirectory","ConnectClientDrivesAtLogon", _
    "ConnectClientPrintersAtLogon","DefaultToMainPrinter")
  strHTML = strHTML & ReadPropertiesSimple("Environment","ADSI Extension for Terminal Services",arrProperties)
  'End all settings on the Environment Properties Page
  
  'All settings on the Sessions Properties Page
  arrProperties = Array("MaxDisconnectionTime","MaxConnectionTime", _
    "MaxIdleTime","BrokenConnectionAction","ReconnectionAction")
  strHTML = strHTML & ReadPropertiesSimple("Sessions","ADSI Extension for Terminal Services",arrProperties)
  'End all settings on the Sessions Properties Page
  
  'All settings on the Remote Control Properties page
  arrProperties = Array("EnableRemoteControl")
  strHTML = strHTML & ReadPropertiesSimple("Remote Control","ADSI Extension for Terminal Services",arrProperties)
  strHTML = strHTML & "Select Case objItem.EnableRemoteControl" & chr(10)
  strHTML = strHTML & "  Case 0" & chr(10)
  strHTML = strHTML & "    WScript.Echo " & chr(34) & _
    "Remote Control disabled" & chr(34) & chr(10)
  strHTML = strHTML & "  Case 1" & chr(10)
  strHTML = strHTML & "    WScript.Echo " & chr(34) & _
    "Remote Control enabled" & chr(34) & chr(10)
  strHTML = strHTML & "    WScript.Echo " & chr(34) & _
    "User permission required" & chr(34) & chr(10)
  strHTML = strHTML & "    WScript.Echo " & chr(34) & _
    "Interact with the session" & chr(34) & chr(10)
  strHTML = strHTML & "  Case 2" & chr(10)
  strHTML = strHTML & "    WScript.Echo " & chr(34) & _
    "Remote Control enabled" & chr(34) & chr(10)
  strHTML = strHTML & "    WScript.Echo " & chr(34) & _
    "User permission not required" & chr(34) & chr(10)
  strHTML = strHTML & "    WScript.Echo " & chr(34) & _
    "Interact with the session" & chr(34) & chr(10)
  strHTML = strHTML & "  Case 3" & chr(10)
  strHTML = strHTML & "    WScript.Echo " & chr(34) & _
    "Remote Control enabled" & chr(34) & chr(10)
  strHTML = strHTML & "    WScript.Echo " & chr(34) & _
    "User permission required" & chr(34) & chr(10)
  strHTML = strHTML & "    WScript.Echo " & chr(34) & _
    "View the session" & chr(34) & chr(10)
  strHTML = strHTML & "  Case 4" & chr(10)
  strHTML = strHTML & "    WScript.Echo " & chr(34) & _
    "Remote Control enabled" & chr(34) & chr(10)
  strHTML = strHTML & "    WScript.Echo " & chr(34) & _
    "User permission not required" & chr(34) & chr(10)
  strHTML = strHTML & "    WScript.Echo " & chr(34) & _
    "View the session" & chr(34) & chr(10)
    
  strHTML = strHTML & "End Select" & chr(10) & chr(10)
  'End all settings on the Remote Control Properties Page
  
  'All settings on the Terminal Services Profile Properties page
  arrProperties = Array("TerminalServicesProfilePath", _
    "TerminalServicesHomeDirectory","TerminalServicesHomeDrive", _
    "AllowLogon")
  strHTML = strHTML & ReadPropertiesSimple("Terminal Services Profile","ADSI Extension for Terminal Services",arrProperties)
  'End all settings on the Terminal Services Profile Properties page
  
  'The attribute on the COM+ Properties page  
  arrSVStringAttribsCOM = Array("msCOM-UserPartitionSetLink")
  strHTML = strHTML & strReadCodeSV("COM+",arrSVStringAttribsCOM)
  'End the attribute on the COM+ Properties page
  
  'All attributes on the Member-Of Properties Page
  arrSVStringAttribsMO = Array("primaryGroupID")
  strHTML = strHTML & strReadCodeSV("Member Of",arrSVStringAttribsMO)
  
  arrMVStringAttribsMO = Array("memberOf")
  strHTML = strHTML & strReadCodeMV("Member Of",arrMVStringAttribsMO)
  'End all attributes on the Member-Of Properties Page
  
  'Selected attributes on the Object Properties Page
  arrSVStringAttribsObj = Array("whenCreated","whenChanged")
  strHTML = strHTML & strReadCodeSV("Object",arrSVStringAttribsObj)
  
  'Added this because canonicalName is an operational attribute
  strHTML = strHTML & "objItem.GetInfoEx Array(" & _
    chr(34) & "canonicalName" & chr(34) & "), 0" & chr(10)
  
  arrMVStringAttribsObj = Array("canonicalName")
  strHTML = strHTML & strReadCodeMV("Object",arrMVStringAttribsObj)
  'End all attributes on the Object Properties Page
  
  'Dial-in and Security pages skipped. 
  'A later version of the tool might include
  'a script to read these Properties pages.
    
  UserAttribsToRead = strHTML
End Function

Function ContactAttribsToRead
  'All attributes on the General Properites Page
  arrSVStringAttribsGP = Array("name", "givenName","initials", _
    "sn","displayName","description","physicalDeliveryOfficeName", _
    "telephoneNumber","mail","wWWHomePage")
  strHTML = strReadCodeSV("General",arrSVStringAttribsGP) 
  
  arrMVStringAttribsGP = Array("otherTelephone", "url")
  strHTML = strHTML & strReadCodeMV("General",arrMVStringAttribsGP)
  'End General Properties Page
  
  'All attributes on the Address Properties Page
  arrSVStringAttribsAP = Array("streetAddress", "l", "st", _
    "postalCode", "c")
  strHTML = strHTML & strReadCodeSV("Address",arrSVStringAttribsAP)

  arrMVStringAttribsAP = Array("postOfficeBox")
  strHTML = strHTML & strReadCodeMV("Address",arrMVStringAttribsAP)
  'End Address Properties Page

  'All attributes on the Telephones Properties Page
  arrSVStringAttribsTele = Array("homePhone","pager", _
    "mobile","facsimileTelephoneNumber","ipPhone", "info")
  strHTML = strHTML & strReadCodeSV("Telephone",arrSVStringAttribsTele)
  
  arrMVStringAttribsTele = Array("otherHomePhone","otherPager", _
    "otherMobile","otherFacsimileTelephoneNumber","otherIpPhone")
  strHTML = strHTML & strReadCodeMV("Telephone",arrMVStringAttribsTele)
  'End Telephones Properties Page 
  
  'All attributes on the Organization Properties Page
  arrSVStringAttribsOrg = Array("title","department", _
    "company","manager")
  strHTML = strHTML & strReadCodeSV("Organization",arrSVStringAttribsOrg)
  
  arrMVStringAttribsOrg = Array("directReports")
  strHTML = strHTML & strReadCodeMV("Organization",arrMVStringAttribsOrg)
  'End Organization Properties Page 
  
  'All attributes on the Member-Of Properties Page
  arrMVStringAttribsMO = Array("memberOf")
  strHTML = strHTML & strReadCodeMV("Member Of",arrMVStringAttribsMO)
  'End all attributes on the Member-Of Properties Page
  
  'Selected attributes on the Object Properties Page
  arrSVStringAttribsObj = Array("whenCreated","whenChanged")
  strHTML = strHTML & strReadCodeSV("Object",arrSVStringAttribsObj)
  
  'Added this because canonicalName is an operational attribute
  strHTML = strHTML & "objItem.GetInfoEx Array(" & _
    chr(34) & "canonicalName" & chr(34) & "), 0" & chr(10)
  
  arrMVStringAttribsObj = Array("canonicalName")
  strHTML = strHTML & strReadCodeMV("Object",arrMVStringAttribsObj)
  'End all attributes on the Object Properties Page
  
  'Security page skipped. 
  'A later version of the tool might include
  'a script to read this Properties page.
    
  ContactAttribsToRead = strHTML
End Function

Function GroupAttribsToRead 
  'All attributes on the General Properties Page
  arrSVStringAttribsGP = Array("name","samAccountName","description", _
    "mail")
  strHTML = strHTML & strReadCodeSV("General",arrSVStringAttribsGP)
  'For reading the bit flags in grouptype
  arrGTConstants = Array("ADS_GROUP_TYPE_GLOBAL_GROUP", _
    "ADS_GROUP_TYPE_DOMAIN_LOCAL_GROUP", _
    "ADS_GROUP_TYPE_UNIVERSAL_GROUP","ADS_GROUP_TYPE_SECURITY_ENABLED")
  arrGTValues = Array("&h2","&h4","&h8","&h80000000")
  strHTML = strHTML & IntReadCode("General","groupType", _
    arrGTConstants,arrGTValues)
  strHTML = strHTML & "If intgroupType AND " & _
   "objHash.Item(" & chr(34) & "ADS_GROUP_TYPE_DOMAIN_LOCAL_GROUP" & chr(34) & ") Then" & chr(10)
  strHTML = strHTML & "  WScript.Echo " & chr(34) & "Group Scope: Domain Local Group" & chr(34) & chr(10)
  strHTML = strHTML & "ElseIf intGroupType AND " & _
   "objHash.Item(" & chr(34) & "ADS_GROUP_TYPE_GLOBAL_GROUP" & chr(34) & ") Then" & chr(10)
  strHTML = strHTML & "  WScript.Echo " & chr(34) & "Group Scope: Global Group" & chr(34) & chr(10)
  strHTML = strHTML & "ElseIf intGroupType AND " & _
   "objHash.Item(" & chr(34) & "ADS_GROUP_TYPE_UNIVERSAL_GROUP" & chr(34) & ") Then" & chr(10)
  strHTML = strHTML & "  WScript.Echo " & chr(34) & "Group Scope: Universal Group" & chr(34) & chr(10) 
  strHTML = strHTML & "End If" & chr(10)  
  strHTML = strHTML & "If intgroupType AND " & _
    "objHash.Item(" & chr(34) & "ADS_GROUP_TYPE_SECURITY_ENABLED" & chr(34) & ") Then" & chr(10)
  strHTML = strHTML & "  WScript.Echo " & chr(34) & "Group Type: Security" & chr(34) & chr(10)
  strHTML = strHTML & "Else" & chr(10)
  strHTML = strHTML & "  WScript.Echo " & chr(34) & "Group Type: Distribution" & chr(34) & chr(10)
  strHTML = strHTML & "End If" & chr(10)
  'End for reading the bit flags in grouptype
  'End General Properties Page
  
  'All attributes on the Managed By page. Code checks to see if the 
  'first field has a value. If so, it binds to the dn of the object (user) 
  'and gets the properties of the object specified on the Managed By page.
  arrSVStringAttribsMB = Array("managedBy")
  strHTML = strHTML & strReadCodeSV("Managed By",arrSVStringAttribsMB)
  strHTML = strHTML & "If strmanagedBy <> " & chr(34) & chr(34) & " Then" & chr(10) 
  strHTML = strHTML & "  Set objItem1 = GetObject(" & chr(34) & "LDAP://" & chr(34) & _
    " &amp; strManagedBy)" & chr(10)
  arrMBProperties = Array("physicalDeliveryOfficeName","streetAddress", _
    "l","c","telephoneNumber","facsimileTelephoneNumber") 
  For Each prop in arrMBProperties
    strHTML = strHTML & "  WScript.Echo " & _
      chr(34) & prop & ": " & chr(34) & " &amp; _" & chr(10) & _
      "    objItem1." & prop & chr(10)
  Next
  strHTML = strHTML & "End If" & chr(10) & chr(10)
   'End all attributes on the Managed By page
  
  'All attributes on the Member Properties Page
  arrMVStringAttribsMO = Array("member")
  strHTML = strHTML & strReadCodeMV("Member",arrMVStringAttribsMO)
  'End all attributes on the Member-Of Properties Page  
  
  'All attributes on the Member-Of Properties Page
  arrMVStringAttribsMO = Array("memberOf")
  strHTML = strHTML & strReadCodeMV("Member Of",arrMVStringAttribsMO)
  'End all attributes on the Member-Of Properties Page
  
  'All attributes on the Managed By page
  'Code checks to see if the 
  'first field has a value. If so, it binds to the dn of 
  'the object (user or group) and gets the properties of the object
  'specified on the Managed By page.
  arrSVStringAttribsMB = Array("managedBy")
  strHTML = strHTML & strReadCodeSV("Managed By",arrSVStringAttribsMB)
  strHTML = strHTML & "If strmanagedBy <> " & chr(34) & chr(34) & " Then" & chr(10) 
  strHTML = strHTML & "  Set objItem1 = GetObject(" & chr(34) & "LDAP://" & chr(34) & _
    " &amp; strManagedBy)" & chr(10)
  arrMBProperties = Array("physicalDeliveryOfficeName","streetAddress", _
    "l","c","telephoneNumber","facsimileTelephoneNumber") 
  For Each prop in arrMBProperties
    strHTML = strHTML & "  WScript.Echo " & _
      chr(34) & prop & ": " & chr(34) & " &amp; _" & chr(10) & _
      "    objItem1." & prop & chr(10)
  Next
  strHTML = strHTML & "End If" & chr(10) & chr(10)
  'End all attributes on the Managed By page
  
  GroupAttribsToRead = strHTML
End Function

Function OUAttribsToRead 
  'All attributes on the General Properties Page
  arrSVStringAttribsGP = Array("name", "description", _
    "streetAddress", "postOfficeBox","l","st","postalCode","c")
  strHTML = strHTML & strReadCodeSV("General",arrSVStringAttribsGP)
  'End General Properties Page
  
  'All attributes on the Managed By page
  'Code checks to see if the first field has a value. If so, it binds to the dn of 
  'the object (user or group) and gets the properties of the object specified on the
  'Managed By page.
  arrSVStringAttribsMB = Array("managedBy")
  strHTML = strHTML & strReadCodeSV("Managed By",arrSVStringAttribsMB)
  strHTML = strHTML & "If strmanagedBy <> " & chr(34) & chr(34) & " Then" & chr(10) 
  strHTML = strHTML & "  Set objItem1 = GetObject(" & chr(34) & "LDAP://" & chr(34) & _
    " &amp; strManagedBy)" & chr(10)
  arrMBProperties = Array("physicalDeliveryOfficeName","streetAddress", _
    "l","c","telephoneNumber","facsimileTelephoneNumber") 
  For Each prop in arrMBProperties
    strHTML = strHTML & "  WScript.Echo " & _
      chr(34) & prop & ": " & chr(34) & " &amp; _" & chr(10) & _
      "    objItem1." & prop & chr(10)
  Next
  strHTML = strHTML & "End If" & chr(10) & chr(10)
   'End all attributes on the Managed By page
  
  'Selected attributes on the Object Properties Page
  arrSVStringAttribsObj = Array("whenCreated","whenChanged")
  strHTML = strHTML & strReadCodeSV("Object",arrSVStringAttribsObj)
  
  'Added this because canonicalName is an operational attribute
  strHTML = strHTML & "objItem.GetInfoEx Array(" & _
    chr(34) & "canonicalName" & chr(34) & "), 0" & chr(10)
  
  arrMVStringAttribsObj = Array("canonicalName")
  strHTML = strHTML & strReadCodeMV("Object",arrMVStringAttribsObj)
  'End all attributes on the Object Properties Page
  
  'Selected attributes on the Group Policy Properties Page
  arrSVStringAttribsGP = Array("gPLink","gPOptions")
  strHTML = strHTML & strReadCodeSV("Group Policy",arrSVStringAttribsGP)
  strHTML = strHTML & "If strgPOptions = 1 Then " & chr(10)
  strHTML = strHTML & "  WScript.Echo " & chr(34) & _
    "Policy inheritance is blocked." & chr(34) & chr(10)
  strHTML = strHTML & "ElseIf strgPOptions = 0 Then " & chr(10)
  strHTML = strHTML & "  WScript.Echo " & chr(34) & _
    "Policies are inherited." & chr(34) & chr(10)
  strHTML = strHTML & "End If" & chr(10) & chr(10)
  'End selected attributes on the Group Policy Properties Page
  
  OUAttribsToRead = strHTML
End Function

Function ComputerAttribsToRead
  'All attributes on the General Properties Page
  arrSVStringAttribsGP = Array("name","dnsHostName","description")
  strHTML = strHTML & strReadCodeSV("General",arrSVStringAttribsGP)
  'End General Properties Page
   
  'For reading the bit flags in userAccountControl
  arrUACConstants = Array("ADS_UF_TRUSTED_FOR_DELEGATION", _
    "ADS_UF_WORKSTATION_TRUST_ACCOUNT","ADS_UF_SERVER_TRUST_ACCOUNT")
  arrUACValues = Array("&h80000","&h1000","&h2000")
  strHTML = strHTML & IntReadCode("General","userAccountControl", _
    arrUACConstants,arrUACValues)
  strHTML = strHTML & "If intuserAccountControl AND objHash.Item(" & chr(34) & "ADS_UF_TRUSTED_FOR_DELEGATION" & chr(34) & ") Then" & chr(10)
  strHTML = strHTML & "  WScript.Echo " & chr(34) & "Trust computer for delegation" & chr(34) & chr(10)
  strHTML = strHTML & "Else" & chr(10)
  strHTML = strHTML & "  WScript.Echo " & chr(34) & "Do not trust the computer for delegation" & chr(34) & chr(10)
  strHTML = strHTML & "End If" & chr(10)  
  strHTML = strHTML & "If intuserAccountControl AND objHash.Item(" & chr(34) & "ADS_UF_SERVER_TRUST_ACCOUNT" & chr(34) & ") Then" & chr(10)
  strHTML = strHTML & "  WScript.Echo " & chr(34) & "Role: Domain Controller" & chr(34) & chr(10)
  strHTML = strHTML & "Else" & chr(10)
  strHTML = strHTML & "  WScript.Echo " & chr(34) & "Role: Workstation or Server" & chr(34) & chr(10)
  strHTML = strHTML & "End If" & chr(10)  
  'End for reading the bit flags in userAccountControl
  
  'All attributes on the Operating System Properties Page
  arrSVStringAttribsGP = Array("operatingSystem","operatingSystemVersion", _
    "operatingSystemServicePack")
  strHTML = strHTML & strReadCodeSV("Operating System",arrSVStringAttribsGP)
  'End Operating System Properties Page
  
  'All attributes on the Member-Of Properties Page
  arrSVStringAttribsMO = Array("primaryGroupID")
  strHTML = strHTML & strReadCodeSV("Member Of",arrSVStringAttribsMO)
  
  arrMVStringAttribsMO = Array("memberOf")
  strHTML = strHTML & strReadCodeMV("Member Of",arrMVStringAttribsMO)
  'End all attributes on the Member-Of Properties Page
  
  'All attributes on the Location Properties Page
  arrSVStringAttribsMO = Array("location")
  strHTML = strHTML & strReadCodeSV("Location",arrSVStringAttribsMO)
  'End all attributes on the Location Properties Page  
    
  'All attributes on the Managed By page. The code checks to see if the 
  'first field has a value. If so, it binds to the dn of the object (user)
  'and gets the properties of the object specified on the Managed By page.
  arrSVStringAttribsMB = Array("managedBy")
  strHTML = strHTML & strReadCodeSV("Managed By",arrSVStringAttribsMB)
  strHTML = strHTML & "If strmanagedBy <> " & chr(34) & chr(34) & " Then" & chr(10) 
  strHTML = strHTML & "  Set objItem1 = GetObject(" & chr(34) & "LDAP://" & chr(34) & _
    " &amp; strManagedBy)" & chr(10)
  arrMBProperties = Array("physicalDeliveryOfficeName","streetAddress", _
    "l","c","telephoneNumber","facsimileTelephoneNumber") 
  For Each prop in arrMBProperties
    strHTML = strHTML & "  WScript.Echo " & _
      chr(34) & prop & ": " & chr(34) & " &amp; _" & chr(10) & _
      "    objItem1." & prop & chr(10)
  Next
  strHTML = strHTML & "End If" & chr(10) & chr(10)
   'End all attributes on the Managed By page
  ComputerAttribsToRead = strHTML
  'Dial-in page skipped. 
  'A later version of the tool might include
  'a script to read this Properties page.
End Function

'***************************************************************************
'* These functions write script code to the code window for the Write 
'* an Object task. Each function varies based on the selected class. 
'* The function name describes the class it supports.
'***************************************************************************
Function UserAttribsToWrite
  'Attributes on the General Properties Page
  arrSVStringAttribsGP = Array("givenName","initials", _
    "sn","displayName","description","physicalDeliveryOfficeName", _
    "telephoneNumber","mail","wWWHomePage")
  strHTML = StrWriteCodeSV("General",arrSVStringAttribsGP,"VALUE")

  arrMVStringAttribsGP = Array("otherTelephone", "url")
  strHTML = strHTML & strWriteCodeMV("General",arrMVStringAttribsGP,"VALUE")
  'End General Properties Page
  
  'Attributes on the Address Properties Page
  arrSVStringAttribsAP = Array("streetAddress", "l", "st", _
    "postalCode")
  strHTML = strHTML & strWriteCodeSV("Address",arrSVStringAttribsAP,"VALUE")
  
  arrSVStringAttribsAP = Array("c")
  strHTML = strHTML & _
    strWriteCodeSV("Address",arrSVStringAttribsAP,"COUNTRY CODE VALUE")

  arrMVStringAttribsAP = Array("postOfficeBox")
  strHTML = strHTML & strWriteCodeMV("Address",arrMVStringAttribsAP,"VALUE")
  'End Address Properties Page
  
  'Attributes on the Profile Properties Page
  arrSVStringAttribsPrP = Array("profilePath", "scriptPath", _
    "homeDirectory")
  strHTML = strHTML & strWriteCodeSV("Profile",arrSVStringAttribsPrP,"VALUE")
  
  arrSVStringAttribsAP = Array("homeDrive")
  strHTML = strHTML & _
    strWriteCodeSV("Profile",arrSVStringAttribsAP,"DRIVE LETTER VALUE (no colon)")
  'End Profile Properties Page
  
  'Attributes on the Telephones Properties Page
  arrSVStringAttribsTele = Array("homePhone","pager", _
    "mobile","facsimileTelephoneNumber","ipPhone", "info")
  strHTML = strHTML & strWriteCodeSV("Telephone",arrSVStringAttribsTele,"VALUE")
  
  arrMVStringAttribsTele = Array("otherHomePhone","otherPager", _
    "otherMobile","otherFacsimileTelephoneNumber","otherIpPhone")
  strHTML = strHTML & strWriteCodeMV("Telephone",arrMVStringAttribsTele,"VALUE")
  'End Telephones Properties Page 
  
  'Attributes on the Organization Properties Page
  arrSVStringAttribsOrg = Array("title","department", _
    "company")
  strHTML = strHTML & strWriteCodeSV("Organization",arrSVStringAttribsOrg,"VALUE")
  
  arrSVStringAttribsOrg = Array("manager")
  strHTML = strHTML & _
    strWriteCodeSV("Organization",arrSVStringAttribsOrg,"DISTINGUISHED NAME VALUE")
 'End attributes on the Organization Properties Page
 
  UserAttribsToWrite = strHTML
  
  'Account, Terminal Services (Remote control, Terminal Services Profile, 
  'Environment, and Sessions) COM+, Dial-in and Security 
  'Properties pages skipped. A later version of the tool might include
  'a script to write some or all these Properties pages.
  'The Member Of properties page contains the memberOf backlink attribute. 
  'Modify the member property of a group to modify the contents 
  'of the memberOf attribute
End Function

Function ContactAttribsToWrite
'Attributes on the General Properties Page
  arrSVStringAttribsGP = Array("givenName","initials", _
    "sn","displayName","description","physicalDeliveryOfficeName", _
    "telephoneNumber","mail","wWWHomePage")
  strHTML = StrWriteCodeSV("General",arrSVStringAttribsGP,"VALUE")

  arrMVStringAttribsGP = Array("otherTelephone", "url")
  strHTML = strHTML & strWriteCodeMV("General",arrMVStringAttribsGP,"VALUE")
  'End General Properties Page
  
  'Attributes on the Address Properties Page
  arrSVStringAttribsAP = Array("streetAddress", "l", "st", _
    "postalCode")
  strHTML = strHTML & strWriteCodeSV("Address",arrSVStringAttribsAP,"VALUE")
  
  arrSVStringAttribsAP = Array("c")
  strHTML = strHTML & _
    strWriteCodeSV("Address",arrSVStringAttribsAP,"COUNTRY CODE VALUE")

  arrMVStringAttribsAP = Array("postOfficeBox")
  strHTML = strHTML & strWriteCodeMV("Address",arrMVStringAttribsAP,"VALUE")
  'End Address Properties Page
  
  'Attributes on the Telephones Properties Page
  arrSVStringAttribsTele = Array("homePhone","pager", _
    "mobile","facsimileTelephoneNumber","ipPhone", "info")
  strHTML = strHTML & strWriteCodeSV("Telephone",arrSVStringAttribsTele,"VALUE")
  
  arrMVStringAttribsTele = Array("otherHomePhone","otherPager", _
    "otherMobile","otherFacsimileTelephoneNumber","otherIpPhone")
  strHTML = strHTML & strWriteCodeMV("Telephone",arrMVStringAttribsTele,"VALUE")
  'End Telephones Properties Page 
  
  'Attributes on the Organization Properties Page
  arrSVStringAttribsOrg = Array("title","department", _
    "company")
  strHTML = strHTML & strWriteCodeSV("Organization",arrSVStringAttribsOrg,"VALUE")
  
  arrSVStringAttribsOrg = Array("manager")
  strHTML = strHTML & _
    strWriteCodeSV("Organization",arrSVStringAttribsOrg,"DISTINGUISHED NAME VALUE")
 'End attributes on the Organization Properties Page

  ContactAttribsToWrite = strHTML
  
  'The Member Of properties page contains the memberOf backlink attribute. 
  'Modify the member property of a group to modify the contents 
  'of the memberOf attribute
End Function

Function GroupAttribsToWrite
  'Selected attributes on the General Properties Page
  arrSVStringAttribsGP = Array("samAccountName","description", _
    "mail")
  strHTML = strHTML & strWriteCodeSV("General",arrSVStringAttribsGP,"VALUE")
  'End attributes on the General Properties Page
  
  'Attributes on the Member Properties Page
  arrMVStringAttribsMO = Array("member")
  strHTML = strHTML & strWriteCodeMV("Member",arrMVStringAttribsMO,"DISTINGUISHED NAME VALUE")
  'End all attributes on the Member-Of Properties Page 
  
  'Attributes on the Managed By Properties Page
  arrSVStringAttribsMB = Array("managedBy")
  strHTML = strHTML & strWriteCodeSV("Managed By",arrSVStringAttribsMB,"DISTINGUISHED NAME VALUE")
  'End attributes on the Managed By Properties Page 
  
  GroupAttribsToWrite = strHTML
  
  'The Member Of properties page contains the memberOf backlink attribute. 
  'Modify the member property of a group to modify the contents 
  'of the memberOf attribute
End Function

Function OUAttribsToWrite
  'Selected attributes on the General Properties Page
  arrSVStringAttribsGP = Array("description", _
    "street", "postOfficeBox","l","st","postalCode")
  strHTML = strHTML & strWriteCodeSV("General",arrSVStringAttribsGP,"VALUE")
  'End General Properties Page
  
  arrSVStringAttribsAP = Array("c")
  strHTML = strHTML & _
    strWriteCodeSV("Address",arrSVStringAttribsAP,"COUNTRY CODE VALUE")
  
  'Attributes on the Managed By Properties Page
  arrSVStringAttribsMB = Array("managedBy")
  strHTML = strHTML & strWriteCodeSV("Managed By",arrSVStringAttribsMB,"DISTINGUISHED NAME VALUE")
  'End attributes on the Managed By Properties Page 
  
  OUAttribsToWrite = strHTML
  
  'COM+, and Group Policy Properties pages skipped. 
  'A later version of the tool might include
  'a script to write some or all these Properties pages.
End Function

Function ComputerAttribsToWrite
  'Selected attributes on the General Properties Page
  arrSVStringAttribsGP = Array("description")
  strHTML = strHTML & strWriteCodeSV("General",arrSVStringAttribsGP,"VALUE")
  'End General Properties Page
  
  'All attributes on the Location Properties Page
  arrSVStringAttribsMO = Array("location")
  strHTML = strHTML & strWriteCodeSV("Location",arrSVStringAttribsMO,"VALUE")
  'End all attributes on the Location Properties Page 
  
  'Attributes on the Managed By Properties Page
  arrSVStringAttribsMB = Array("managedBy")
  strHTML = strHTML & strWriteCodeSV("Managed By",arrSVStringAttribsMB,"DISTINGUISHED NAME VALUE")
  'End attributes on the Managed By Properties Page 
  
  ComputerAttribsToWrite = strHTML
  
  'The Operating System properties page contains attributes that are written when
  'a computer becomes a member of the domain.
  'The Member Of properties page contains the memberOf backlink attribute. 
  'Modify the member property of a group to modify the contents 
  'of the memberOf attribute.
  
  'Dial-in page skipped. A later version of the tool might include
  'a script to read this Properties page.
End Function

'***************************************************************************
'* These functions manipulate the script code that appears in the code
'* window. Function details appear above each function.
'***************************************************************************

'Reformat the class name so that the first character of the class
'name is uppercase. This does not have 
'an impact on the script's ability to run properly.
Function ReformatObjName
  strChar1 = UCase(Left(oFormADScripto.ClassesPulldown.Value,1))
  strRemaining = LCase(Mid(oFormADScripto.ClassesPulldown.Value,2))
  ReformatObjName = "obj" & strChar1 & strRemaining
End Function

'Determine whether the naming attribute for a
'method should be ou or cn
Function NamingAttribute
  If oFormADScripto.ClassesPulldown.Value = "organizationalUnit" Then
    NamingAttribute = "ou"
  Else
    NamingAttribute = "cn"
  End If
End Function

'Size the code window and write some header script code based on the 
'selected task.
Function PreAmble(intCols,intRows)
  strHTML = "<textarea name=codeText cols=" & intCols & " rows=" & intRows & " style=""width:90%;border-width:0;"">"
  If oFormADScripto.TaskSelectPulldown.Value = "Create an Object" AND oFormADScripto.ClassesPullDown.Value = "group" Then
    arrGroupEnum = Array("ADS_GROUP_TYPE_GLOBAL_GROUP = &h2", _
     "ADS_GROUP_TYPE_LOCAL_GROUP = &h4","ADS_GROUP_TYPE_UNIVERSAL_GROUP = &h8", _
     "ADS_GROUP_TYPE_SECURITY_ENABLED = &h80000000")
    For Each Constant In arrGroupEnum
      strHTML = strHTML & chr(10) & Constant
    Next
    strHTML = strHTML & chr(10) & chr(10) 
  End If
  strHTML = strHTML & "strContainer = " & chr(34) & "" & chr(34) & chr(10)
  
  strHTML = strHTML & "strName = " & chr(34) & "EzAd" & _
    UCase(Left(oFormADScripto.ClassesPullDown.Value,1)) & _
    Mid(oFormADScripto.ClassesPullDown.Value,2) & chr(34) & chr(10)
  If oFormADScripto.TaskSelectPulldown.Value = "Write an Object" AND oFormADScripto.ClassesPullDown.Value <> "computer" Then
    strHTML = strHTML & chr(10) & "Const ADS_PROPERTY_CLEAR = 1" & chr(10) 
    strHTML = strHTML & "Const ADS_PROPERTY_UPDATE = 2" & chr(10) 
    strHTML = strHTML & "Const ADS_PROPERTY_APPEND = 3" & chr(10) 
    strHTML = strHTML & "Const ADS_PROPERTY_DELETE = 4"
  ElseIf oFormADScripto.TaskSelectPulldown.Value = "Write an Object" AND oFormADScripto.ClassesPullDown.Value = "computer" Then
    strHTML = strHTML & "'This page contains only single-valued attributes" & chr(10)
    strHTML = strHTML & "'so only the ADS_PROPERTY_CLEAR constant is defined here." & chr(10)
    strHTML = strHTML & chr(10) & "Const ADS_PROPERTY_CLEAR = 1" 
  End If
  PreAmble = strHTML & chr(10)
End Function

'Generate the binding string text for the script
Function BindString
  If oFormADScripto.TaskSelectPulldown.Value = "Create an Object" OR _
    oFormADScripto.TaskSelectPulldown.Value = "Delete an Object" Then
    strObj = "objContainer"
    strHTML = strHTML & "'***********************************************" & chr(10)
    strHTML = strHTML & "'*         Connect to a container              *" & chr(10) 
    strHTML = strHTML & "'***********************************************" & chr(10)
    
    strHTML = strHTML & "Set objRootDSE = GetObject(" & _
      chr(34) & "LDAP://rootDSE" & chr(34) & ")" & chr(10)
    strHTML = strHTML & "If strContainer = " & chr(34) & chr(34) & " Then" & chr(10)
      strHTML = strHTML & "  Set " & strObj & " = GetObject(" & chr(34) & "LDAP://" & chr(34) & _
        " &amp; _" & chr(10) & _
          "    objRootDSE.Get(" & chr(34) & "defaultNamingContext" & chr(34) & "))" & chr(10)
    strHTML = strHTML & "Else" & chr(10)
      strHTML = strHTML & "  Set " & strObj &  " = GetObject(" & chr(34) & "LDAP://" & chr(34) & _
        " &amp; strContainer &amp; " & chr(34) & "," & chr(34) & " &amp; _" & chr(10) & _
          "    objRootDSE.Get(" & chr(34) & "defaultNamingContext" & chr(34) & "))" & chr(10)
    strHTML = strHTML & "End If" & chr(10)
    'The remarked section adds error testing to determine if the attempted 
    'connection to a domain failed.
    'strHTML = strHTML & "If err.number = 424 Then" & chr(10)
    'strHTML = strHTML & "  WScript.Echo " & chr(34) & "You must run the script from an Active Directory" & _
    '  " enabled client." & chr(34) & chr(10)
    'strHTML = strHTML & "  WScript.Quit" & chr(10)
    'strHTML = strHTML & "End If" & chr(10)
    'strHTML = strHTML & "On Error GoTo 0" & chr(10)
    strHTML = strHTML & "'***********************************************" & chr(10)
    strHTML = strHTML & "'*       End connect to a container            *" & chr(10) 
    strHTML = strHTML & "'***********************************************" & chr(10) & chr(10)
  Else
    strObj = "objItem"
    strNamingAttribute = NamingAttribute()
    strHTML = strHTML & "'***********************************************" & chr(10)
    strHTML = strHTML & "'*          Connect to an object                 *" & chr(10) 
    strHTML = strHTML & "'***********************************************" & chr(10)
    strHTML = strHTML & "Set objRootDSE = GetObject(" & _
      chr(34) & "LDAP://rootDSE" & chr(34) & ")" & chr(10)
    strHTML = strHTML & "If strContainer = " & chr(34) & chr(34) & " Then" & chr(10)
    If oFormADScripto.ClassesPulldown.Value = "organizationalUnit" Then
      strHTML = strHTML & "  arrNameExceptions = Array(" & chr(34) & "Users" & chr(34) & _
        "," & chr(34) & "Computers" & chr(34) & _
        "," & chr(34) & "Builtin" & chr(34) & _
        "," & chr(34) & "System" & chr(34) & _
        ", _" & chr(10) & _
        "    " & chr(34) & "ForeignSecurityPrincipals" & chr(34) & _
        "," & chr(34) & "LostAndFound" & chr(34) & ")" & chr(10)
      strHTML = strHTML & "  For Each name in arrNameExceptions" & chr(10)
      strHTML = strHTML & "    If lcase(strName) = lcase(name) Then" & chr(10)
      strHTML = strHTML & "      strNameAttrib = " & chr(34) & "cn=" & chr(34) & chr(10)
      strHTML = strHTML & "      Exit For" & chr(10)
      strHTML = strHTML & "    Else" & chr(10)
      strHTML = strHTML & "      strNameAttrib = " & chr(34) & "ou=" & chr(34) & chr(10)
      strHTML = strHTML & "    End If" & chr(10)
      strHTML = strHTML & "  Next" & chr(10)
    strHTML = strHTML & "  Set " & strObj & " = GetObject(" & chr(34) & "LDAP://" & chr(34) & _
      " &amp; strNameAttrib &amp; strName &amp; " & chr(34) & "," & chr(34) & " &amp; _" & chr(10) & _
         "    objRootDSE.Get(" & chr(34) & "defaultNamingContext" & chr(34) & "))" & chr(10)
    strHTML = strHTML & "Else" & chr(10)
      strHTML = strHTML & "  Set " & strObj &  " = GetObject(" & chr(34) & "LDAP://" & _
        strNamingAttribute & "=" & chr(34) & " &amp; strName &amp; " & _
          chr(34) & "," & chr(34) & " &amp; strContainer &amp; " & chr(34) & "," & chr(34) & " &amp; _" & chr(10) & _
          "    objRootDSE.Get(" & chr(34) & "defaultNamingContext" & chr(34) & "))" & chr(10)
    Else
      strHTML = strHTML & "  Set " & strObj & " = GetObject(" & chr(34) & "LDAP://" & chr(34) & _
      " &amp; _" & chr(10) & _
         "    objRootDSE.Get(" & chr(34) & "defaultNamingContext" & chr(34) & "))" & chr(10)
      strHTML = strHTML & "Else" & chr(10)
      strHTML = strHTML & "  Set " & strObj &  " = GetObject(" & chr(34) & "LDAP://" & _
        strNamingAttribute & "=" & chr(34) & " &amp; strName &amp; " & _
          chr(34) & "," & chr(34) & " &amp; strContainer &amp; " & chr(34) & "," & chr(34) & " &amp; _" & chr(10) & _
          "    objRootDSE.Get(" & chr(34) & "defaultNamingContext" & chr(34) & "))" & chr(10)
    End If
    strHTML = strHTML & "End If" & chr(10)
    'The remarked section adds error testing to determine if the attempted 
    'connection to a domain failed.
    'strHTML = strHTML & "If err.number = 424 Then" & chr(10)
    'strHTML = strHTML & "  WScript.Echo " & chr(34) & "You must run the script from an Active Directory" & _
    '  " enabled client." & chr(34) & chr(10)
    'strHTML = strHTML & "  WScript.Quit" & chr(10)
    'strHTML = strHTML & "End If" & chr(10)
    strHTML = strHTML & "'***********************************************" & chr(10)
    strHTML = strHTML & "'*         End connect to an object           *" & chr(10) 
    strHTML = strHTML & "'***********************************************" & chr(10) & chr(10)
  End If
  BindString = strHTML
End Function

'***************************************************************************
'* These two routines generate the text that appears in the two modal dialog
'* boxes. These dialog boxes appear when theh Running This Script or
'* Script Notes buttons are pressed.
'***************************************************************************
Sub RunDialog()
  'Create a file system object to populate the modal dialog box.
  Set objFileSystem = CreateObject("Scripting.FileSystemObject")
  'Create an HTML file named tempEZscripto.htm
  Set objFile = objFileSystem.CreateTextFile("EZscriptoRunNotes.htm",True)
  objFile.WriteLine("<html>")
  objFile.WriteLine("<title>Before Running The " & _
    oFormADScripto.TaskSelectPulldown.Value & " - " & _
    UCase(Left(oFormADScripto.ClassesPulldown.Value,1)) & _
    Mid(oFormADScripto.ClassesPullDown.Value,2) &  " Script</title>")
  objFile.WriteLine("<head>")
  objFile.WriteLine("<style>")
  objFile.WriteLine("BODY{background-color: beige;font-family: arial;font-size:10pt;margin-left:10px;}")
  objFile.WriteLine("div.head{font-size:12pt;font-weight:bold;}")
  objFile.WriteLine("div.code{font-size:10pt;font-family:courier;margin-left:10px}")  
  objFile.WriteLine("UL{margin-top:5px;margin-bottom:5px;}")
  objFile.WriteLine("</style>")
  objFile.WriteLine("</head>")
  objFile.WriteLine("<body>")
  Select Case oFormADScripto.TaskSelectPulldown.Value
    Case "Create an Object"
      arrAltText = Array("create",lcase(oFormADScripto.TaskSelectPulldown.Value)," to create")
    Case "Write an Object"
      arrAltText = Array("write",lcase(oFormADScripto.TaskSelectPulldown.Value)," whose attributes you will assign")
      strAddlNotes = "<div class = head>Attribute Values</div>" & _
        "<UL><li><i>VALUE</i> or <i>VALUEn</i> - string values" & _
        "<li><i>COUNTRY CODE VALUE</i> - This value is a two-digit country code. For a " & _
        "list of country codes, see <A href=http://www.iso.org>The ISO Web site</A>" & _
        "<li><i>DRIVE LETTER VALUE</i> - This value is a drive letter, " & _
        "typically a value between F and Z. Do not specify a colon following the letter." & _
        "<li><i>DISTINGUISHED NAME VALUE</i> - This value is the DN of an object.<br>" & _
        "<b>Examples</b><br>" & _
        "The MyerKen user account in the Management OU of the NA.fabrikam.com domain:" & _ 
        "<div class=code>cn=myerken,ou=management,dc=na,dc=fabrikam,dc=com</div>" & _ 
        "The Atl-Users group in the Users container of the contoso.com domain:" & _ 
        "<div class=code>cn=atl-users,cn=users,dc=contoso,dc=com</div>" & _ 
        "</UL>"
    Case "Read an Object"
      arrAltText = Array("read",lcase(oFormADScripto.TaskSelectPulldown.Value)," whose attributes you will read")
    Case "Delete an Object"
      arrAltText = Array("delete",lcase(oFormADScripto.TaskSelectPulldown.Value)," to delete")  
  End Select

  strNotes = _
    "<div class=head>Script Operation</div>"
    Select Case oFormADScripto.TaskSelectPulldown.value
      Case "Create an Object"
        strAction = "creates a"
      Case "Delete an Object"
        strAction = "deletes a "
      Case "Write an Object"
        strAction = "writes the attributes of the"
      Case "Read an Object"
        strAction = "reads the attributes of the"
    End Select
      
    strNotes = strNotes & _
      "This script, " & strAction & " " & oFormADScripto.ClassesPulldown.value & _
      " whose name is specified by the strName variable and whose location" & _
      " is specified by the strContainer variable.<br>" & _
      " NOTE: A default strName variable of EzAd" & _
        UCase(Left(oFormADScripto.ClassesPullDown.Value,1)) & _
        Mid(oFormADScripto.ClassesPullDown.Value,2) & _
      " appears in the script.<br><br>" & _
      "<div class=head>The strContainer Variable</div>" & _
      "Add a value (between quotes) for strContainer " & _
      "on the first script line with the RDN of the " & _
      "container where the " & _
      oFormADScripto.ClassesPulldown.Value & arrAltText(2) & " is located.<br>" & _ 
      "<b>Examples</b><br>" & _
      "To " & arrAltText(1) & " in the Users container:" & _
      "<div class=code>strContainer = " & chr(34) & "cn=Users" & chr(34) & "</div>" & _
      "To " & arrAltText(1) & " in the Finance OU:" & _
      "<div class=code>strContainer = " & chr(34) & "ou=Finance" & chr(34) & "</div>" & _
      "To " & arrAltText(1) & " in the HR OU below the Depts OU<br>" & _
      "<div class=code>strContainer = " & chr(34) & "ou=HR,ou=Depts" & chr(34) & "</div><br>" & _
      "<b>Important:</b>" & _
      "<UL><li>If you plan to " & arrAltText(1) & " in a child container of another container, specify " & _
      "the child OU before the parent OU, as the previous example demonstrates.<br>" & _
      "<li>Do not specify a value for strContainer if you want to " & arrAltText(1) & _
      " directly below the domain container.</UL><br>" & _
      "<div class=head>The strName Variable</div>" & _
      "Add a value (between quotes) for strName on the second script line with " & _
      "the name of the object you want to " & arrAltText(0) & ".<br>" & _
      "<b>Examples:</b><br>" & _
      "To " & arrAltText(1) & " named HR:" & _
      "<div class=code>strName = " & chr(34) & "HR" & chr(34) & "</div>" & _
      "To " & arrAltText(1) & " named MyerKen:" & _
      "<div class=code>strName = " & chr(34) & "myerken" & chr(34) & "</div><br>" & _
      "Note, do not specify the naming attribute, i.e., CN or OU.<br><br>"
      If oFormADScripto.ClassesPullDown.Value = "group" Then
        strNotes = strNotes & "<div class=head>The groupType attribute</div>" & _
        "If the groupType attribute is not specified when you create a group, Active Directory " & _
        "creates a security group with global scope. This is typically referred to as simply a global " & _
        "global group. However, by changing the constants appearing with the groupType attribute in " & _
        "the script, you can create any allowed group types. The following list shows how the script code for the " & _
        "groupType attribute should appear in the script for each type of Active Directory group:<br><br>" & _
        "To create a distribution group with global scope:" & _
        "<div class=code>objGroup.Put " & chr(34) & "groupType" & chr(34) & ", ADS_GROUP_TYPE_GLOBAL_GROUP</div>" & _
        "To create a distribution group with local scope:" & _
        "<div class=code>objGroup.Put " & chr(34) & "groupType" & chr(34) & ", ADS_GROUP_TYPE_LOCAL_GROUP</div>" & _
        "To create a distribution group with universal scope:" & _
        "<div class=code>objGroup.Put " & chr(34) & "groupType" & chr(34) & ", ADS_GROUP_TYPE_UNIVERSAL_GROUP</div>" & _
        "To create a security group with global scope:" & _
        "<div class=code>objGroup.Put " & chr(34) & "groupType" & chr(34) & ", ADS_GROUP_TYPE_GLOBAL_GROUP Or _<br>" & _
          "&nbsp;&nbsp;ADS_GROUP_TYPE_SECURITY_ENABLED</div>" & _
        "To create a security group with local scope:" & _
        "<div class=code>objGroup.Put " & chr(34) & "groupType" & chr(34) & ", ADS_GROUP_TYPE_LOCAL_GROUP Or _<br>" & _
          "&nbsp;&nbsp;ADS_GROUP_TYPE_SECURITY_ENABLED</div>" & _
        "To create a security group with universal scope:" & _
        "<div class=code>objGroup.Put " & chr(34) & "groupType" & chr(34) & ", ADS_GROUP_TYPE_UNIVERSAL_GROUP Or _<br>" & _
          "&nbsp;&nbsp;ADS_GROUP_TYPE_SECURITY_ENABLED</div><br><br>"
      End If
      strNotes = strNotes & "<div class=head>Running This Script</div>" & _
      "To run this script, make sure that your workstation is connected to an " & _
      "Active Directory domain and ADSI is installed. ADSI is installed by default " & _
      "on computers running Windows 2000, Windows XP, and all members of the Windows .NET Server family " & _
      "of operating systems.<br><br>"
    
  'Write the contents of strNotes to the file object
  objFile.WriteLine(strNotes)
  'This variable specifies the height of the modal dialog box to hold the 
  'contents of strNotes without requiring the operator to scroll the dialog
  'box.
  If oFormADScripto.TaskSelectPulldown.Value = "Write an Object" Then
    objFile.WriteLine(strAddlNotes)
    intHeight = 700
  Else
    intHeight = 600
  End If
  
  objFile.WriteLine("</body>")
  objFile.WriteLine("</html>")
  objFile.Close

  'Use a modal dialog box so that the operator must close the window before
  'returning to the main window.
  window.showModalDialog "EZscriptoRunNotes.htm",, _
    "dialogHeight:" & intHeight & "px; dialogWidth:550px; edge:raised; help:No; resizable:no;"
  Set objFile = objFileSystem.GetFile("EZscriptoRunNotes.htm")
  objFile.Delete
End Sub

Sub ImptDialog()
  'Create a file system object to populate the modal dialog box.
  Set objFileSystem = CreateObject("Scripting.FileSystemObject")
  'Create an HTML file named tempEZscripto.htm
  Set objFile = objFileSystem.CreateTextFile("EZscriptoNotes.htm",True)
  objFile.WriteLine("<html>")
  objFile.WriteLine("<title>Important Notes About This Script</title>")
  objFile.WriteLine("<head>")
  objFile.WriteLine("<style>")
  objFile.WriteLine("BODY{background-color: beige;font-family: arial;font-size:10pt;margin-left:10px;}")
  objFile.WriteLine("div.head{font-size:12pt;font-weight:bold;}")
  objFile.WriteLine("div.code{font-size:10pt;font-family:courier;margin-left:10px}")
  objFile.WriteLine("UL{margin-top:5px;margin-bottom:5px;}")
  objFile.WriteLine("</style>")
  objFile.WriteLine("</head>")
  objFile.WriteLine("<body>")
  Select Case oFormADScripto.TaskSelectPulldown.Value
    Case "Write an Object"
      strNotes = "<div class=head>Modifying The Script</div>" & _
      "<ul><li>Replace <i>VALUE</i> with the value you want to write" & _
      " to each attribute." & _
      "<li>Remove the entire line that starts with: <b>objItem.Put ...</b>, for any attribute " & _
      " you don't want to write.</ul>" & _
      "<b>Note:</b> A subset of all " & oFormADScripto.ClassesPulldown.Value & " attributes appear in the" & _
      " scripts.<br><br>" & _
      "<div class=head>The PutEx Constants (ADS Property Enumeration)</div>" & _
      "The PutEx method with the ADS_PROPERTY_UPDATE constant specified in the script, replaces" & _
      " the entries stored in an attribute." & _
      " You can change this behavior by specifying ADS_PROPERTY_APPEND to add entries, or " & _
      " ADS_PROPERTY_DELETE to delete entries.<br> " & _
      " You can also use the ADS_PROPERTY_CLEAR constant to delete all entries from a" & _
      " single or multivalued attribute. To use" & _
      " ADS_PROPERTY_CLEAR you must replace <i>VALUEx</i> with 0.<br>" & _
      " The following example clears the description single-valued attribute:<br>" & _
      "<div class=code>objItem.PutEx ADS_PROPERTY_CLEAR " & chr(34) & "description" & chr(34) & ", 0</div>" & _
      " The following example clears the otherTelephone multivalued attribute:<br>" & _
      "<div class=code>objItem.PutEx ADS_PROPERTY_CLEAR " & chr(34) & "otherTelephone" & chr(34) & ", 0</div><br>" & _
      "See the <i>Microsoft System Administration Scripting Resource Kit</i> for more information."
      'Write the contents of strNotes to the file object
      objFile.WriteLine(strNotes)
      'This variable specifies the height of the modal dialog box to hold the 
      'contents of strNotes without requiring the operator to scroll the dialog
      'box.
      intHeight = 380
    Case "Delete an Object"
      If oFormADScripto.ClassesPulldown.Value = "organizationalUnit" Then
        strNotes = "A container must be empty before you can delete " & _
          "it using the Delete method. See the <i>ADSI SDK</i> for information " & _
          "on the <b>IADsDeleteOps</b> interface for deleting a non-empty container."
          objFile.WriteLine(strNotes)
        intHeight = 100
      End If
    Case "Read an Object"
      If oFormADScripto.ClassesPulldown.Value = "user" Then
        strNotes = "The Terminal Services Extension interface " & _
          "used to read the settings on the <b>Environment, Sessions, " & _
          "Remote Control</b>, and <b>Terminal Services Profile</b> pages is " & _
          "available from a Windows XP client or any member of the " & _
          "Windows .NET Server family of operating systems.<br><br>" & _
          "See the <i>Microsoft System Administration Scripting Resource Kit</i> for more information."
        objFile.WriteLine(strNotes)
        intHeight = 140
      ElseIf oFormADScripto.ClassesPulldown.Value = "organizationalUnit" Then
        strNotes = "These OUs contain only the description attribute on the" & _
          " General Properties page and no Group Policy tab: <b>Users, Computers, Builtin, System, ForeignSecurityPrincipals,</b> " & _
          "and <b>LostAndFound</b>"
        objFile.WriteLine(strNotes)
        intHeight = 100
      End If
  End Select
  
  objFile.WriteLine("</body>")
  objFile.WriteLine("</html>")
  objFile.Close
  
  window.showModalDialog "EZscriptoNotes.htm",, _
    "dialogHeight:" & intHeight & "px; dialogWidth:550px; edge:raised; help:No; resizable:no;"
  Set objFile = objFileSystem.GetFile("EZscriptoNotes.htm")
  objFile.Delete
End Sub

'***************************************************************************
'* These routines generate the script that appears in the code window.
'* In each case, the code checks to see if the operator has selected 
'* a class to complete a task. If so, the code enables buttons
'* and calls functions that generate script code.
'***************************************************************************

Sub CreateCreateScript
  If oFormADScripto.ClassesPulldown.Value = "PulldownMessage" Then
    oFormADScripto.notes_button.value = ""
    oFormAction.action(1).disabled = True
    oFormAction.action(0).disabled = True
    Exit Sub
  End If
  
  strNamingAttribute = NamingAttribute()
  strHTML = PreAmble(87,23)
  strHTML = strHTML & BindString()
  strObjName = ReformatObjName()
  
  strHTML = strHTML & "Set " & strObjName & " = " & _
    "objContainer.Create(" & chr(34) & oFormADScripto.ClassesPulldown.Value & chr(34) & _
      ", " & chr(34) & strNamingAttribute & "=" & chr(34) & " &amp; strName)" & chr(10)
   
  If oFormADScripto.ClassesPulldown.Value = "user" Or oFormADScripto.ClassesPulldown.Value = "group" Then
    strHTML = strHTML & strObjName & ".Put " & chr(34) & _
        "sAMAccountName" & chr(34) & ", " & "strName" & chr(10)
  End If
  
  If oFormADScripto.ClassesPullDown.Value = "group" Then
    strHTML = strHTML & strObjName & ".Put " & chr(34) & _
        "groupType" & chr(34) & ", " & "ADS_GROUP_TYPE_GLOBAL_GROUP Or _" & chr(10) & _
        "  ADS_GROUP_TYPE_SECURITY_ENABLED" & chr(10)
  End If
  
  strHTML = strHTML & strObjName & ".SetInfo"
  oDivADScriptCode.InnerHTML = strHTML
  oFormADScripto.codeText.Wrap = "off"
    
  'Enable the Running This Script, Run, and Save buttons. 
  Call FinalUIState
End Sub

Sub CreateDeleteScript
  If oFormADScripto.ClassesPulldown.Value = "PulldownMessage" Then
    oFormADScripto.notes_button.value =  ""
    oFormAction.action(1).disabled = True
    oFormAction.action(0).disabled = True
    Exit Sub
  End If
  
  strObjName = ReformatObjName()
  strHTML = PreAmble(87,21)
  strHTML = strHTML & BindString()
  
  If oFormADScripto.ClassesPulldown.Value = "organizationalUnit" Then
    oFormADScripto.notes_button.value = "Click here for important information about" & _
      " deleting non-empty containers"
  End If
    
  strNamingAttribute = NamingAttribute()
  
  strHTML = strHTML & "objContainer.Delete " & chr(34) & _
    oFormADScripto.ClassesPulldown.Value & chr(34) & _
      ", " & chr(34) & strNamingAttribute & "=" & chr(34) & " &amp; strName" & chr(10)
  
  oDivADScriptCode.InnerHTML = strHTML
  oFormADScripto.codeText.Wrap = "off"
  Call FinalUIState
End Sub

Sub CreateReadScript
  
  If oFormADScripto.ClassesPulldown.Value = "PulldownMessage" Then
    oFormADScripto.notes_button.value=""
    oFormAction.action(1).disabled = True
    oFormAction.action(0).disabled = True
    Exit Sub
  End If
  If oFormADScripto.ClassesPulldown.Value = "organizationalUnit" Then
    oFormADScripto.notes_button.value = "Click here for important information about" & _
      " containers with a limited set of OU attributes"
  End If
  
  strObjName = ReformatObjName()
  strHTML = strHTML & PreAmble(87,25)
  strHTML = strHTML & "On Error Resume Next" & Chr(10) & Chr(10)
  strHTML = strHTML & BindString()
  
  Select Case oFormADScripto.ClassesPulldown.Value
    Case "user"
      strHTML = strHTML & UserAttribsToRead()
      oDivADScriptCode.InnerHTML = strHTML
      oFormADScripto.notes_button.value = "Click here for important information about" & _
      " reading Terminal Services settings"
    Case "group"
      strHTML = strHTML & GroupAttribsToRead()
      oDivADScriptCode.InnerHTML = strHTML
    Case "contact"
      strHTML = strHTML & ContactAttribsToRead()
      oDivADScriptCode.InnerHTML = strHTML
    Case "organizationalUnit"
      strHTML = strHTML & OUAttribsToRead()
      oDivADScriptCode.InnerHTML = strHTML
    Case "computer"
      strHTML = strHTML & ComputerAttribsToRead()
      oDivADScriptCode.InnerHTML = strHTML
    Case Else
      oDivADScriptCode.InnerHTML = ""
  End Select
  
  oFormADScripto.codeText.Wrap = "off"
  
  Call FinalUIState

End Sub

Sub CreateWriteScript
  
  If oFormADScripto.ClassesPulldown.Value = "PulldownMessage" Then
    oFormADScripto.notes_button.value=""
    oFormAction.action(1).disabled = True
    oFormAction.action(0).disabled = True
    Exit Sub
  End If
  
  oFormADScripto.notes_button.value = "Click here for important information about modifying write scripts"
  strObjName = ReformatObjName()
  strHTML = strHTML & PreAmble(87,25)
  strHTML = strHTML & Chr(10)
  strHTML = strHTML & BindString()
  
  Select Case oFormADScripto.ClassesPulldown.Value
    Case "user"
      strHTML = strHTML & UserAttribsToWrite()
      oDivADScriptCode.InnerHTML = strHTML
    Case "group"
      strHTML = strHTML & GroupAttribsToWrite()
      oDivADScriptCode.InnerHTML = strHTML
    Case "contact"
      strHTML = strHTML & ContactAttribsToWrite()
      oDivADScriptCode.InnerHTML = strHTML
    Case "organizationalUnit"
      strHTML = strHTML & OUAttribsToWrite()
      oDivADScriptCode.InnerHTML = strHTML
    Case "computer"
      strHTML = strHTML & ComputerAttribsToWrite()
      oDivADScriptCode.InnerHTML = strHTML
    Case Else
      oDivADScriptCode.InnerHTML = ""
  End Select
  
  oDivADScriptCode.InnerHTML = strHTML & "objItem.SetInfo"  
  oFormADScripto.codeText.Wrap = "off"
  
  Call FinalUIState

End Sub

'****************************************************************************
'* When the operator presses the Run button, we use the WshShell object's Run
'* method to run the code currently in the textarea under cscript.exe. we use
'* cmd.exe's /k parameter to ensure the command window remains visible after
'* the script has finished running.
'****************************************************************************

Sub RunScript
   Set objFS = CreateObject("Scripting.FileSystemObject")
   strTmpName = "temp_script.vbs"
   Set objScript = objFS.CreateTextFile(strTmpName)
   objScript.Write oDivADScriptCode.InnerText
   objScript.Close
   Set objShell = CreateObject("WScript.Shell")
   strCmdLine = "cmd /k cscript.exe "
   strCmdLine = strCmdLine & strTmpName
   objShell.Run(strCmdLine)
End Sub

'****************************************************************************
'* When the operator presses the Save button, we present them with an InputBox
'* and force them to give us the full path to where they'd like to the save
'* the script that is currently in the textarea. The user is probably quite
'* upset with our laziness here....and who can blame them?
'****************************************************************************

Sub SaveScript
   Set objFSO = CreateObject("Scripting.FileSystemObject")
   strSaveFileName = InputBox("Please enter the complete path where you want to save your script (for example, C:\Scripts\MyScript.vbs).")
   If strSaveFileName = "" Then
      Exit Sub
   End If
   Set objFile = objFSO.CreateTextFile(strSaveFileName)
   objFile.WriteLine oDivADScriptCode.InnerText
   objFile.Close
End Sub

'****************************************************************************
'* When the operator presses the Open button, we present them with an InputBox
'* and force them to give us the full path to the script they'd like to open.
'* This is, of course, rather wonky - but it's meant to be.
'****************************************************************************

Sub OpenScript
   Set objFSO = CreateObject("Scripting.FileSystemObject")
   strOpenFileName = InputBox("Please enter the complete path name for your script (for example, C:\Scripts\MyScript.vbs).")
   If strOpenFileName = "" Then
      Exit Sub
   End If
   Set objFile = objFSO.OpenTextFile(strOpenFileName)
   strHTML = "<textarea cols=100 rows=30>"
   strHTML = strHTML & objFile.ReadAll()
   strHTML = strHTML & "</textarea>"
   oDivADScriptCode.InnerHTML =  strHTML
   objFile.Close
   oFormAction.action(0).disabled = False
   oFormAction.action(1).disabled = False
End Sub

'****************************************************************************
'* When the operator presses the Quit button, the file where we've been storing
'* the scripts gets deleted and the main window closes. 
'****************************************************************************

Sub QuitScript

   On Error Resume Next
   Set objFSO = CreateObject("Scripting.FileSystemObject")
   objFSO.DeleteFile "temp_script.vbs"
   Set objFSO = Nothing
   self.Close()

End Sub