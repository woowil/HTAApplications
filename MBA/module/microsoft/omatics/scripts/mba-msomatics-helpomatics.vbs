Sub RunScript
End Sub

Sub TestSub
    strCopy = oFormHelpomatic.ConfigureArea.Value
    document.parentwindow.clipboardData.SetData "text", strCopy
End Sub

Sub CopyCode
    strCopy = oFormHelpomatic.SubroutineArea.Value
    document.parentwindow.clipboardData.SetData "text", strCopy
End Sub

Sub GenerateCode
    oFormHelpomatic.show_example_button.Disabled = False
    oFormHelpomatic.copy_html_button.Disabled = False
    oFormHelpomatic.reset_button.Disabled = False
    oFormHelpomatic.copy_code_button.Disabled = False

    If oFormHelpomatic.hta_elements.Value = "Autorun Script" Then
        strScript = ""
        oFormHelpomatic.show_example_button.Disabled = True
        oFormHelpomatic.copy_html_button.Disabled = True
        oFormHelpomatic.reset_button.Disabled = True
        strSub = "'Displays a message box when the application loads" & vbcrlf & vbcrlf
        strSub = strSub & "Sub Window_Onload" & vbcrlf
        strSub = strSub & "    Msgbox " & chr(34) & "The application has started." & chr(34) & vbcrlf
        strSub = strSub & "End Sub"
        strExample = strScript
        strDescription = "An autorun script is a script that runs each time the HTA is started or reloaded. (For information on how to reload an HTA, see the element <b>Refresh the HTA</b>.) There is nothing special about an autorun script; it's simply script code contained in a subroutine named <b>Window_Onload</b>. <br>&nbsp;<br>The sample code shown here displays a message box each time the HTA is started or reloaded."
    End If

    If oFormHelpomatic.hta_elements.Value = "Basic HTML tags" Then
        oFormHelpomatic.reset_button.Disabled = True
        oFormHelpomatic.copy_code_button.Disabled = True
        strScript = "<b>boldface text</b> <i>italic text</i> <u>underlined text</u><br>regular script <sub>sub</sub>script <sup>super</sup>script<br>This &nbsp;&nbsp;&nbsp;&nbsp; text &nbsp;&nbsp;&nbsp;&nbsp; is &nbsp;&nbsp;&nbsp;&nbsp; separated &nbsp;&nbsp;&nbsp;&nbsp; by &nbsp;&nbsp;&nbsp;&nbsp; blank &nbsp;&nbsp;&nbsp;&nbsp; spaces."
        strExample = strScript
        strDescription = "This sample illustrates some basic HTML tags, including <b>b</b> (bold), <b>i</b> (italics), and <b>u</b> (underline). These tags help you create formatted output."
    End If

    If oFormHelpomatic.hta_elements.Value="Button" Then
        strScript = "<input id=runbutton  class=" & chr(34) & "button" & chr(34) & " type=" & chr(34) & "button" & chr(34) & " value=" & chr(34) & "Run Button" & chr(34) & " name=" & chr(34) & "run_button" & chr(34) & "  onClick=" & chr(34) & "RunScript" & chr(34) & ">"
        strSub = "'Displays a message box when the button is clicked" & vbcrlf & vbcrlf & "Sub RunScript" & vbcrlf & "    Msgbox " & chr(34) & "Test" & chr(34) & vbcrlf & "End Sub"
        strExample = strScript
        strDescription = "Buttons are used to initiate actions, which, in an HTA, typically means running a script of some kind. With a button, you will primarily be concerned with the <b>Value</b> (which is actually the on-screen label for the button) and the <b>onClick event</b>, with which you specify the script that runs when the button is clicked.<br>&nbsp;<br>In the example, the button simply displays a message box when clicked."
    End If

    If oFormHelpomatic.hta_elements.Value="Checkbox" Then
        strScript = "<input type=" & chr(34) & "checkbox" & chr(34) & "name=" & chr(34) & "BasicCheckbox" & chr(34) & "value=" & chr(34) & "1" & chr(34) & "> Option 1"
        strSub = "'Tells you whether BasicCheckbox has been selected or not" & vbcrlf & vbcrlf  
        strSub = strSub & "Sub RunScript" & vbcrlf
        strSub = strSub & "    If BasicCheckbox.Checked Then" & vbcrlf
        strSub = strSub & "        Msgbox " & chr(34) & "The checkbox has been checked." & chr(34) & vbcrlf
        strSub = strSub & "    Else" & vbcrlf
        strSub = strSub & "        Msgbox " & chr(34) & "The checkbox has not been checked." & chr(34) & vbcrlf
        strSub = strSub & "    End If" & vbcrlf
        strSub = strSub & "End Sub"
        strExample = strScript & "&nbsp;&nbsp;&nbsp;<input id=runbutton  class=" & chr(34) & "button" & chr(34) & " type=" & chr(34) & "button" & chr(34) & " value=" & chr(34) & "Run Button" & chr(34) & " name=" & chr(34) & "run_button" & chr(34) & "  onClick=" & chr(34) & "RunScript" & chr(34) & ">"
        strDescription = "A checkbox is typically used for optional elements within a script: yes, we want to check available memory on a computer; no, we don't want to check available memory on a computer. Interesting things you can do with a checkbox include disabling it (<b>Disabled = True</b>) and configuring it to be automatically selected when displayed (<b>Checked = True</b>).<br>&nbsp;<br>In this example, select or de-select the checkbox and then click the Run Button; the script will determine whether the checkbox has been selected or not, and then inform you of its current state."
    End If

    If oFormHelpomatic.hta_elements.Value = "Data display in new window" Then
        strScript = ""
        oFormHelpomatic.show_example_button.Disabled = True
        oFormHelpomatic.copy_html_button.Disabled = True       
        strSub = "'Displays formatted information in a new window" & vbcrlf & vbcrlf
        strSub = strSub & "Sub RunScript" & vbcrlf
        strSub = strSub & "    Set objIE = CreateObject(" & chr(34) & "InternetExplorer.Application" & chr(34) & ")" & vbcrlf
        strSub = strSub & "    objIE.Navigate(" & chr(34) & "about:blank" & chr(34) & ")" & vbcrlf
        strSub = strSub & "    objIE.ToolBar = 0" & vbcrlf
        strSub = strSub & "    objIE.StatusBar = 0" & vbcrlf
        strSub = strSub & "    Set objDoc = objIE.Document.Body" & vbcrlf
        strSub = strSub & "    strHTML = " & chr(34) & "<B>This information is displayed in a separate window.</B>" & chr(34) & vbcrlf
        strSub = strSub & "    objDoc.InnerHTML = strHTML" & vbcrlf
        strSub = strSub & "    objIE.Visible = True" & vbcrlf
        strSub = strSub & "End Sub"           
        strExample = "<input id=runbutton  class=" & chr(34) & "button" & chr(34) & " type=" & chr(34) & "button" & chr(34) & " value=" & chr(34) & "Run Button" & chr(34) & " name=" & chr(34) & "run_button" & chr(34) & "  onClick=" & chr(34) & "RunScript" & chr(34) & ">"
        strDescription = "By creating a new instance of Internet Explorer, you can display data in a window separate from your main HTA window. To do this, create the new instance of Internet Explorer, and Navigate to about:blank (which causes Internet Explorer to open up with a blank page). You can then write data to this new window by setting the <b>InnerHTML property</b> of the document to the desired values.<br>&nbsp;<br>This example opens a new window and simply displays one formatted sentence.<br>&nbsp;<br><b>Note</b>: Be sure you set the <b>Visible</b> property of the new Internet Explorer instance to True. Otherwise, your window will not be visible on screen."
    End If

    If oFormHelpomatic.hta_elements.Value = "Change cursor" Then
        strScript = ""
        oFormHelpomatic.show_example_button.Disabled = True
        oFormHelpomatic.copy_html_button.Disabled = True       
        strSub = "'Changes the cursor to an hourglass, pauses, then changes it back to the default" & vbcrlf & vbcrlf
        strSub = strSub & "Sub RunScript" & vbcrlf
        strSub = strSub & "    document.body.style.cursor = " & chr(34) & "wait" & chr(34) & vbcrlf
        strSub = strSub & "End Sub"
        strExample = "<input id=runbutton  class=" & chr(34) & "button" & chr(34) & " type=" & chr(34) & "button" & chr(34) & " value=" & chr(34) & "Run Button" & chr(34) & " name=" & chr(34) & "run_button" & chr(34) & "  onClick=" & chr(34) & "RunScript" & chr(34) & ">"
        strDescription = "Using a script, you can modify the cursor to reflect the current state of your HTA; for example, if the HTA is carrying out a lengthy procedure, you can change the cursor to an hourglass. The cursor can be changed to any of the following: crosshair; default; hand; help; wait.<br>&nbsp;<br>In this example, clicking the Run Button will change the cursor to an hourglass. To change it back, modify the script by setting the cursor to <b>default</b>. Click the Reset Example button to apply the new script code, and then click the Run Button again."
    End If

    If oFormHelpomatic.hta_elements.Value = "Copy to Clipboard" Then
        oFormHelpomatic.show_example_button.Disabled = True
        strScript = "<textarea name=" & chr(34) & "BasicTextArea" & chr(34) & " rows=" & chr(34) & "5" & chr(34) & " cols=" & chr(34) & "75" & chr(34) & "></textarea>"
        strSub = "' Copies the contents of the text area to the clipboard" & vbcrlf & vbcrlf
        strSub = strSub & "Sub RunScript" & vbcrlf
        strSub = strSub & "    strCopy = BasicTextArea.Value" & vbcrlf
        strSub = strSub & "    document.parentwindow.clipboardData.SetData " & chr(34) & "text" & chr(34) & ", strCopy" & vbcrlf
        strSub = strSub & "End Sub"
        strExample = strScript & "&nbsp;&nbsp;&nbsp;<input id=runbutton  class=" & chr(34) & "button" & chr(34) & " type=" & chr(34) & "button" & chr(34) & " value=" & chr(34) & "Run Button" & chr(34) & " name=" & chr(34) & "run_button" & chr(34) & "  onClick=" & chr(34) & "RunScript" & chr(34) & ">"
        strDescription = "Sample script that copies the contents of the text area named BasicTextArea to the Clipboard. To test the script, type some information into the text area, and then click the Run Button. The information you typed will now be copied to the Clipboard. (You don't need to select the text; just click the button.)"
    End If

    If oFormHelpomatic.hta_elements.Value = "Create basic HTA" Then
        oFormHelpomatic.reset_button.Disabled = True
        oFormHelpomatic.copy_code_button.Disabled = True
        oFormHelpomatic.show_example_button.Disabled = True
        strScript = "<html>" & vbcrlf
        strScript = strScript & "<head>" & vbcrlf
        strScript = strScript & "<title>HTA Helpomatic</title>" & vbcrlf & vbcrlf
        strScript = strScript & "<HTA:APPLICATION " & vbcrlf
        strScript = strScript & "     ID=" & chr(34) & "objHTAHelpomatic" & chr(34) & vbcrlf
        strScript = strScript & "     APPLICATIONNAME=" & chr(34) & "HTAHelpomatic" & chr(34) & vbcrlf
        strScript = strScript & "     SCROLL=" & chr(34) & "yes" & chr(34) & vbcrlf
        strScript = strScript & "     SINGLEINSTANCE=" & chr(34) & "yes" & chr(34) & vbcrlf
        strScript = strScript & "     WINDOWSTATE=" & chr(34) & "maximize" & chr(34) & vbcrlf
        strScript = strScript & ">" & vbcrlf
        strScript = strScript & "</head>" & vbcrlf & vbcrlf
        strScript = strScript & "<SCRIPT Language=" & chr(34) & "VBScript" & chr(34) & ">" & vbcrlf
        strScript = strScript & "</" & "SCRIPT>" & vbcrlf
        strScript = strScript & "<body>" & vbcrlf & vbcrlf & vbcrlf
        strScript = strScript & "</body>" & vbcrlf
        strScript = strScript & "</html>"
        strDescription = "Sample code that creates the basic structure of an HTA; to use this, copy the code, paste it into Notepad, and then save the file with a .hta file extension. You can then place HTML elements (such as buttons and listboxes) inside the body tag, and place scripts that you write inside the script tag.<br>&nbsp;<br>Note: Be sure your scripts have unique names within each HTA, and that they start with the <b>Sub</b> tag (for example, <b>Sub RunScript</b>) and end with the <b>End Sub</b> tag."
    End If

    If oFormHelpomatic.hta_elements.Value = "Disabled control" Then
        strScript = "<select size=" & chr(34) & "3" & chr(34) &  " name=" & chr(34) & "Dropdown1" & chr(34) &  " Disabled=True>" & vbCrLf
        strScript = strScript & "<option value=" & chr(34) & "1" & chr(34) & ">Option 1</option>" & vbCrLf
        strScript = strScript & "<option value=" & chr(34) & "2" & chr(34) & ">Option 2</option>" & vbCrLf
        strScript = strScript & "<option value=" & chr(34) & "3" & chr(34) & ">Option 3</option>" & vbCrLf
        strScript = strScript & "<option value=" & chr(34) & "4" & chr(34) & ">Option 4</option>" & vbCrLf
        strScript = strScript & "</select>"
        strSub = "'Enables the dropdownl list named DropDown1" & vbcrlf & vbcrlf
        strSub = strSub & "Sub RunScript" & vbcrlf
        strSub = strSub & "    DropDown1.Disabled = False" & vbcrlf
        strSub = strSub & "End Sub"
        strExample = strScript & "&nbsp;&nbsp;&nbsp;<input id=runbutton  class=" & chr(34) & "button" & chr(34) & " type=" & chr(34) & "button" & chr(34) & " value=" & chr(34) & "Run Button" & chr(34) & " name=" & chr(34) & "run_button" & chr(34) & "  onClick=" & chr(34) & "RunScript" & chr(34) & ">"
        strDescription = "Among other things, disabling a control makes that control available only when it makes sense for it to be available. (For example, you might keep a Create New User button disabled until the user has entered all the requisite information, such as user name, password, etc.) You can programmatically enable a disabled control (or vice-versa) by setting the item's Disabled property to False.<br>&nbsp;<br>In this example, clicking the Run Button will enable the listbox. To disable the listbox, change the code so that the <b>Disabled</b> property is True, click the Reset Example button, and then click the Run Button again."
    End If

    If oFormHelpomatic.hta_elements.Value = "Display a confirmation box" Then
        strScript = ""
        oFormHelpomatic.show_example_button.Disabled = True
        oFormHelpomatic.copy_html_button.Disabled = True
        strSub = "'Displays a confirmation box before proceeding with the script" & vbcrlf & vbcrlf
        strSub = strSub & "Sub RunScript" & vbcrlf
        strSub = strSub & "    blnAnswer = window.confirm(" & chr(34) & "Are you sure you want to continue?" & chr(34) & ")" & vbcrlf
        strSub = strSub & "    If blnAnswer Then" & vbcrlf
        strSub = strSub & "        Msgbox " & chr(34) & "You clicked the OK button" & chr(34) & vbcrlf
        strSub = strSub & "    Else" & vbcrlf
        strSub = strSub & "        Msgbox " & chr(34) & "You clicked the Cancel button." & chr(34) & vbcrlf
        strSub = strSub & "    End If" & vbcrlf
        strSub = strSub & "End Sub" & vbcrlf
        strExample = "<input id=runbutton  class=" & chr(34) & "button" & chr(34) & " type=" & chr(34) & "button" & chr(34) & " value=" & chr(34) & "Run Button" & chr(34) & " name=" & chr(34) & "run_button" & chr(34) & "  onClick=" & chr(34) & "RunScript" & chr(34) & ">"
        strDescription = "The Confirm method displays a dialog box that contains both an OK and a Cancel button. If the return value from the confirmation box is True, that means the user clicked the OK button; if the return value is False, that means the user clicked the Cancel button.<br>&nbsp;<br>The sample script displays a message box asking, 'Are you sure you want to continue?' To display a different message, simply set that message as the sole parameter passed to the Confirm method."
    End If

    If oFormHelpomatic.hta_elements.Value = "Display the print dialog box" Then
        strScript = ""
        oFormHelpomatic.show_example_button.Disabled = True
        oFormHelpomatic.copy_html_button.Disabled = True
        oFormHelpomatic.reset_button.Disabled = True
        strSub = "'Display the print dialog box" & vbcrlf & vbcrlf
        strSub = strSub & "Sub RunScript" & vbcrlf
        strSub = strSub & "    Window.Print()" & vbcrlf
        strSub = strSub & "End Sub"
        strExample = "<input id=runbutton  class=" & chr(34) & "button" & chr(34) & " type=" & chr(34) & "button" & chr(34) & " value=" & chr(34) & "Run Button" & chr(34) & " name=" & chr(34) & "run_button" & chr(34) & "  onClick=" & chr(34) & "RunScript" & chr(34) & ">"
        strDescription = "This code displays the Print dialog box, enabling you to print the current page. Alternatively, you can right-click anywhere in the document and select Print to display this dialog box. A button such as this can be particularly useful in cases where you are displaying data in a new (and presumably uncluttered) window."
    End If

    If oFormHelpomatic.hta_elements.Value = "Drop-down Listbox" Then
        strScript = "<select size=" & chr(34) & "1" & chr(34) &  " name=" & chr(34) & "DropDown1" & chr(34) &  " onChange=" & chr(34) & "RunScript" & chr(34) & ">" & vbCrLf
        strScript = strScript & "<option value=" & chr(34) & "1" & chr(34) & ">Option 1</option>" & vbCrLf
        strScript = strScript & "<option value=" & chr(34) & "2" & chr(34) & ">Option 2</option>" & vbCrLf
        strScript = strScript & "<option value=" & chr(34) & "3" & chr(34) & ">Option 3</option>" & vbCrLf
        strScript = strScript & "<option value=" & chr(34) & "4" & chr(34) & ">Option 4</option>" & vbCrLf
        strScript = strScript & "</select>"
        strSub = "'Indicates which option in DropDown1 was selected" & vbcrlf & vbcrlf  
        strSub = strSub & "Sub RunScript" & vbcrlf
        strSub = strSub & "    Msgbox DropDown1.Value" & vbcrlf
        strSub = strSub & "End Sub" & vbcrlf
        strExample = strScript 
        strDescription = "A drop-down listbox allows you to select one item from any number of options. The HTML code for a drop-down listbox consists of a select tag (which creates the list) and a number of option tags, with each option tag representing one item in the drop-down list. Each item in the list should have a unique value; to set a particular option as the default value, set its <b>Selected</b> property to True.<br>&nbsp;<br>The sample drop-down listbox shows four options; when you select one, a script (RunScript) runs, and a message box apears displaying the value of the selected item."
    End If

    If oFormHelpomatic.hta_elements.Value = "Font Formatting" Then
        oFormHelpomatic.reset_button.Disabled = True
        oFormHelpomatic.copy_code_button.Disabled = True
        strScript = "<font color=" & chr(34) & "red" & chr(34) & " face=" & chr(34) & "Times New Roman" & chr(34) & " size=" & chr(34) & "6" & chr(34) & ">Your text goes here</font>"
        strExample = strScript
        strDescription = "HTML gives you a considerable amount of control over fonts and font formatting. This example uses basic HTML tags to set the font color (red), typeface (Times New Roman) and size (6, based on an HTML standard where 1 is the smallest font and 6 the largest). Font formats can also be controlled in an HTA by using styles and stylesheets."
    End If

    If oFormHelpomatic.hta_elements.Value = "Hyperlink" Then
        oFormHelpomatic.reset_button.Disabled = True
        oFormHelpomatic.copy_code_button.Disabled = True
        strScript = "<a href=" & chr(34) & "http://www.microsoft.com/technet/scriptcenter" & chr(34) & " target=" & chr(34) & "blank" & chr(34) & ">Script Center</a href>"
        strExample = strScript
        strDescription = "Hyperlinks enable you to open Web pages (either from the local computer or across the Internet) and display them. The sample HTML code shown here opens the Script Center home page in a separate window (<b>target= 'blank'</b>)."
    End If

    If oFormHelpomatic.hta_elements.Value = "Listbox" Then
        strScript = "<select size=" & chr(34) & "3" & chr(34) &  " name=" & chr(34) & "Listbox1" & chr(34) &  " onChange=" & chr(34) & "RunScript" & chr(34) & ">" & vbCrLf
        strScript = strScript & "<option value=" & chr(34) & "1" & chr(34) & ">Option 1</option>" & vbCrLf
        strScript = strScript & "<option value=" & chr(34) & "2" & chr(34) & ">Option 2</option>" & vbCrLf
        strScript = strScript & "<option value=" & chr(34) & "3" & chr(34) & ">Option 3</option>" & vbCrLf
        strScript = strScript & "<option value=" & chr(34) & "4" & chr(34) & ">Option 4</option>" & vbCrLf
        strScript = strScript & "</select>"
        strSub = "'Indicates which option in Listbox1 was selected" & vbcrlf & vbcrlf  
        strSub = strSub & "Sub RunScript" & vbcrlf
        strSub = strSub & "    Msgbox Listbox1.Value" & vbcrlf
        strSub = strSub & "End Sub" 
        strExample = strScript 
        strDescription = "A listbox is nothing more than a drop-down listbox with a <b>Size</b> larger than 1. (The size indicates how many items are displayed on screen at one time.)<br>&nbsp;<br>The sample code simply displays the listbox item that is currently selected."
    End If

    If oFormHelpomatic.hta_elements.Value = "Multi-select List Box" Then
        strScript = "<select size=" & chr(34) & "3" & chr(34) &  " name=" & chr(34) & "DropDown1" & chr(34) &  " multiple>" & vbCrLf
        strScript = strScript & "<option value=" & chr(34) & "1" & chr(34) & ">Option 1</option>" & vbCrLf
        strScript = strScript & "<option value=" & chr(34) & "2" & chr(34) & ">Option 2</option>" & vbCrLf
        strScript = strScript & "<option value=" & chr(34) & "3" & chr(34) & ">Option 3</option>" & vbCrLf
        strScript = strScript & "<option value=" & chr(34) & "4" & chr(34) & ">Option 4</option>" & vbCrLf
        strScript = strScript & "</select>"
        strSub = "'Displays all the items selected in a multi-select listbox" & vbcrlf & vbcrlf
        strSub = strSub & "Sub RunScript" & vbcrlf
        strSub = strSub & "    For i = 0 to (Dropdown1.Options.Length - 1)" & vbcrlf
        strSub = strSub & "        If (Dropdown1.Options(i).Selected) Then" & vbcrlf
        strSub = strSub & "            strComputer = strComputer & Dropdown1.Options(i).Value & vbcrlf" & vbcrlf
        strSub = strSub & "        End If" & vbcrlf
        strSub = strSub & "    Next" & vbcrlf
        strSub = strSub & "    Msgbox strComputer" & vbcrlf
        strSub = strSub & "End Sub  "
        strExample = strScript & "&nbsp;&nbsp;&nbsp;<input id=runbutton  class=" & chr(34) & "button" & chr(34) & " type=" & chr(34) & "button" & chr(34) & " value=" & chr(34) & "Run Button" & chr(34) & " name=" & chr(34) & "run_button" & chr(34) & "  onClick=" & chr(34) & "RunScript" & chr(34) & ">"
        strDescription = "A multi-select listbox allows you to select more than one item at a time; to do so, click on the first item, then hold down the CTRL key and click on any subsequent items. Any listbox can be turned into a multi-select listbox simply by adding the <b>multiple</b> parameter.<br>&nbsp;<br>The sample script shows how to loop through the items in the listbox and identify which ones have been selected."
    End If

    If oFormHelpomatic.hta_elements.Value = "On-the-fly listbox" Then
        strScript = ""
        oFormHelpomatic.show_example_button.Disabled = True
        oFormHelpomatic.copy_html_button.Disabled = True        
        oFormHelpomatic.reset_button.Disabled = True
        strSub = "'Creates an on-the-fly listbox based on the contents of the text file computers.txt" & vbcrlf & vbcrlf
        strSub = strSub & "Sub Window_Onload" & vbcrlf
        strSub = strSub & "    ForReading = 1" & vbcrlf
        strSub = strSub & "    strNewFile = " & chr(34) & "computers.txt" & chr(34) & vbcrlf
        strSub = strSub & "    Set objFSO = CreateObject(" & chr(34) & "Scripting.FileSystemObject" & chr(34) & ")" & vbcrlf
        strSub = strSub & "    Set objFile = objFSO.OpenTextFile _" & vbcrlf
        strSub = strSub & "        (strNewFile, ForReading)" & vbcrlf
        strSub = strSub & "    Do Until objFile.AtEndOfStream" & vbcrlf
        strSub = strSub & "        strLine = objFile.ReadLine" & vbcrlf
        strSub = strSub & "        Set objOption = Document.createElement(" & chr(34) & "OPTION" & chr(34) & ")" & vbcrlf
        strSub = strSub & "        objOption.Text = strLine" & vbcrlf
        strSub = strSub & "        objOption.Value = strLine" & vbcrlf
        strSub = strSub & "        AvailableComputers.Add(objOption)" & vbcrlf
        strSub = strSub & "    Loop" & vbcrlf
        strSub = strSub & "    objFile.Close" & vbcrlf
        strSub = strSub & "End Sub"   
        strExample = strScript
        strDescription = "Sample code that reads in a list of computer names from a text file (computers.txt) and then constructs a listbox using the names read in from the list. This simple example does not allow you to modify the list after it has been created; to get a different set of computers in the listbox, you will have to edit the computers.txt file and then re-run the code."
    End If

    If oFormHelpomatic.hta_elements.Value="Password box" Then
        strScript = "<input type=" & chr(34) & "password" & chr(34) & " name=" & chr(34) & "PasswordArea" & chr(34) & " size=" & chr(34) & "30" & chr(34) & ">"
        strSub = "'Displays the value typed into the password box" & vbcrlf & vbcrlf
        strSub = strSub & "Sub RunScript" & vbcrlf
        strSub = strSub & "    Msgbox PasswordArea.Value" & vbcrlf
        strSub = strSub & "End Sub"
        strExample = strScript & "&nbsp;&nbsp;&nbsp;<input id=runbutton  class=" & chr(34) & "button" & chr(34) & " type=" & chr(34) & "button" & chr(34) & " value=" & chr(34) & "Run Button" & chr(34) & " name=" & chr(34) & "run_button" & chr(34) & "  onClick=" & chr(34) & "RunScript" & chr(34) & ">"
        strDescription = "A password box masks any data entered into it, displaying black circles rather than the actual characters. However, the actual characters are still recorded by and available to your script. This allows you to enter sensitive information, such as passwords, without having that information displayed on screen. One interesting property you can set here: <b>MaxLength</b>, which limits the number of characters that can be typed into the box.<br>&nbsp;<br>In the example shown here, the information you type in the password box is displayed in a message box when you click Run Button. This simply demonstrates that the masked information is available for use by the script."
    End If

    If oFormHelpomatic.hta_elements.Value = "Paste from Clipboard" Then
        oFormHelpomatic.show_example_button.Disabled = True
        strScript = "<span id=DataArea>This is a span named DataArea.</span>"
        strSub = "' Pastes the contents of the clipboard to the span named DataArea" & vbcrlf & vbcrlf
        strSub = strSub & "Sub RunScript" & vbcrlf
        strSub = strSub & "    DataArea.InnerHTML = document.parentwindow.clipboardData.GetData(" & chr(34) & "text" & chr(34) & ")" & vbcrlf
        strSub = strSub & "End Sub"
        strExample = strScript & "<br>&nbsp;<br><input id=runbutton  class=" & chr(34) & "button" & chr(34) & " type=" & chr(34) & "button" & chr(34) & " value=" & chr(34) & "Run Button" & chr(34) & " name=" & chr(34) & "run_button" & chr(34) & "  onClick=" & chr(34) & "RunScript" & chr(34) & ">"
        strDescription = "The clipboardData.GetData method allows you to programmtically paste text currently on the Clipboard. In a textbox or textarea, you can paste from the Clipboard by typing CTRL-V or by right-clicking and choosing Paste from the context menu. The clipboardData.GetData method, however, enables you to paste Clipboard text into other parts of your document, such as a span.<br>&nbsp;<br>To use the sample code, copy some text to the Clipboard, then click the Run Button to insert that text into the span named DataArea."
    End If

    If oFormHelpomatic.hta_elements.Value = "Prompt the user for input" Then
        oFormHelpomatic.show_example_button.Disabled = True
        oFormHelpomatic.copy_html_button.Disabled = True
        strSub = "'Displays a dialog box prompting the user for input" & vbcrlf & vbcrlf
        strSub = strSub & "Sub RunScript" & vbcrlf
        strSub = strSub & "    strAnswer = window.prompt(" & chr(34) & "Please enter the domain name." & chr(34) & ", " & chr(34) & "fabrikam.com" & chr(34) & ")" & vbcrlf
        strSub = strSub & "    If IsNull(strAnswer) Then" & vbcrlf
        strSub = strSub & "        Msgbox " & chr(34) & "You clicked the Cancel button" & chr(34) & vbcrlf
        strSub = strSub & "    Else" & vbcrlf
        strSub = strSub & "        Msgbox " & chr(34) & "You entered: " & chr(34) & " & strAnswer" & vbcrlf
        strSub = strSub & "    End If" & vbcrlf
        strSub = strSub & "End Sub"
        strExample = "<input id=runbutton  class=" & chr(34) & "button" & chr(34) & " type=" & chr(34) & "button" & chr(34) & " value=" & chr(34) & "Run Button" & chr(34) & " name=" & chr(34) & "run_button" & chr(34) & "  onClick=" & chr(34) & "RunScript" & chr(34) & ">"
        strDescription = "This sample script displays a dialog box, complete with OK and Cancel buttons, that prompts the user to enter data of some kind. If the return value for the input box is null, that means the user clicked the Cancel button; if the user clicks OK, the return value will be whatever has been typed into the input box. (If the user does not type anything yet still clicks OK, the return value will be an empty string rather than null.) If you click the Run Button and respond to the input box, the sample script will report the return value.<br>&nbsp;<br>The input box accepts two parameters: the prompt message and, optionally, a default value."
    End If

    If oFormHelpomatic.hta_elements.Value="Radio button" Then
        strScript = "<input type=" & chr(34) & "radio" & chr(34) & " name=" & chr(34) & "UserOption" & chr(34) & " value=" & chr(34) & "1" & chr(34) & ">Option 1<BR>" & vbCrLf
        strScript = strScript & "<input type=" & chr(34) & "radio" & chr(34) & " name=" & chr(34) & "UserOption" & chr(34) & " value=" & chr(34) & "2" & chr(34) & ">Option 2<BR>" & vbCrLf
        strScript = strScript & "<input type=" & chr(34) & "radio" & chr(34) & " name=" & chr(34) & "UserOption" & chr(34) & " value=" & chr(34) & "3" & chr(34) & ">Option 3<BR>" & vbCrLf
        strSub = "'Indicates which radio button was selected" & vbcrlf & vbcrlf
        strSub = strSub & "Sub RunScript" & vbcrlf
        strSub = strSub & "    If UserOption(0).Checked Then" & vbcrlf
        strSub = strSub & "        Msgbox " & chr(34) & "Option 1 was selected." & chr(34) & vbcrlf
        strSub = strSub & "    End If" & vbcrlf
        strSub = strSub & "    If UserOption(1).Checked Then" & vbcrlf
        strSub = strSub & "        Msgbox " & chr(34) & "Option 2 was selected." & chr(34) & vbcrlf
        strSub = strSub & "    End If" & vbcrlf
        strSub = strSub & "    If UserOption(2).Checked Then" & vbcrlf
        strSub = strSub & "        Msgbox " & chr(34) & "Option 3 was selected." & chr(34) & vbcrlf
        strSub = strSub & "    End If" & vbcrlf
        strSub = strSub & "End Sub"
        strExample = strScript & "<br><input id=runbutton  class=" & chr(34) & "button" & chr(34) & " type=" & chr(34) & "button" & chr(34) & " value=" & chr(34) & "Run Button" & chr(34) & " name=" & chr(34) & "run_button" & chr(34) & "  onClick=" & chr(34) & "RunScript" & chr(34) & ">"
        strDescription = "Radio buttons are used for choices that are mutually exclusive: you can either disable or enable a user account, but you can't do both at the same time. Radio buttons are meaningless by themselves: you must have at least two of them, and they must have the same Name. If you prefer one of those buttons to represent the default, simply set its <b>Checked</b> property to True.<br>&nbsp;<br>To run the sample script, select a button and then click the Run Button; the script will tell you which option was selected."
    End If

    If oFormHelpomatic.hta_elements.Value = "Refresh the HTA" Then
        oFormHelpomatic.show_example_button.Disabled = True
        oFormHelpomatic.copy_html_button.Disabled = True
        strSub = "'Refreshes the HTA page, which includes re-running any Windows_Onload code" & vbcrlf & vbcrlf
        strSub = strSub & "Sub RunScript" & vbcrlf
        strSub = strSub & "    Location.Reload(True)" & vbcrlf
        strSub = strSub & "End Sub"
        strExample = "<input id=runbutton  class=" & chr(34) & "button" & chr(34) & " type=" & chr(34) & "button" & chr(34) & " value=" & chr(34) & "Run Button" & chr(34) & " name=" & chr(34) & "run_button" & chr(34) & "  onClick=" & chr(34) & "RunScript" & chr(34) & ">"
        strDescription = "Sample code that refreshes the HTA window. In effect, this reloads and restarts the HTA. The Reload method is useful if certain data has changed since the HTA was started. For example, suppose your HTA reads in a list of computer names from a text file, and you have now added some additional names to that file. To display those names, you can simply reload the HTA, which, among other things, will cause it to re-read the text file."
    End If

    If oFormHelpomatic.hta_elements.Value = "Save data" Then
        oFormHelpomatic.show_example_button.Disabled = True
        oFormHelpomatic.copy_html_button.Disabled = True
        strScript = "<span id=DataArea><B>Span</B>. The contents of this span (DataArea) will be saved as a .htm file.</span>"
        strSub = "'Data in the span DataArea is saved to a .htm file" & vbcrlf & vbcrlf
        strSub = strSub & "Sub RunScript" & vbcrlf
        strSub = strSub & "    Set objFSO = CreateObject(" & chr(34) & "Scripting.FileSystemObject" & chr(34) & ")" & vbcrlf
        strSub = strSub & "    objFSO.CreateTextFile(" & chr(34) & "test.htm" & chr(34) & ")" & vbcrlf
        strSub = strSub & "    Set objFile = objFSO.OpenTextFile(" & chr(34) & "test.htm" & chr(34) & ", 2)" & vbcrlf
        strSub = strSub & "    objFile.WriteLine DataArea.InnerHTML" & vbcrlf
        strSub = strSub & "    objFile.Close" & vbcrlf
        strSub = strSub & "End Sub"         
        strExample = strScript & "<br>&nbsp;<br><input id=runbutton  class=" & chr(34) & "button" & chr(34) & " type=" & chr(34) & "button" & chr(34) & " value=" & chr(34) & "Run Button" & chr(34) & " name=" & chr(34) & "run_button" & chr(34) & "  onClick=" & chr(34) & "RunScript" & chr(34) & ">"
        strDescription = "Sample script that saves the information (including all HTML formatting) found in a span named DataArea. In this example, the FileSystemObject is used to create a new file named test.htm, and then the data is saved to that file."

    End If

    If oFormHelpomatic.hta_elements.Value = "Span" Then
        strScript = "<span id=DataArea>This is a span named DataArea.</span>"
        strSub = "'Sets the InnerHTML for the span named DataArea" & vbcrlf & vbcrlf  
        strSub = strSub & "Sub RunScript" & vbcrlf
        strSub = strSub & "    DataArea.InnerHTML = " & chr(34) & "<B>The computer did not respond when pinged.</B>" & chr(34) & vbcrlf
        strSub = strSub & "End Sub"        
        strExample = strScript & "<br>&nbsp;<br><input id=runbutton  class=" & chr(34) & "button" & chr(34) & " type=" & chr(34) & "button" & chr(34) & " value=" & chr(34) & "Run Button" & chr(34) & " name=" & chr(34) & "run_button" & chr(34) & "  onClick=" & chr(34) & "RunScript" & chr(34) & ">"
        strDescription = "The Span tag allows you to name a portion of your document, and then use scripts to programmatically modify that part of the document. A typical use of a span is to: <br>&nbsp;<br>1) Run a script (such as a WMI script) that collects information about something. <br>&nbsp;<br>2) Store the output of the script, along with any HTML formatting you wish to add, to a variable. <br>&nbsp;<br>3) Set the InnerHTML property of the span to that variable (e.g., DataArea.InnerHTML = strHTML). <br>&nbsp;<br>The simple example shown here sets the InnerHTML to a boldfaced sentence, combining the boldface tag with regular text."
    End If

    If oFormHelpomatic.hta_elements.Value = "Timer" Then
        strScript = ""
        oFormHelpomatic.show_example_button.Disabled = True
        oFormHelpomatic.copy_html_button.Disabled = True
        oFormHelpomatic.reset_button.Disabled = True       
        strSub = "'Causes the script RunScript to run every 5 seconds (5000 milliseconds)" & vbcrlf & vbcrlf
        strSub = strSub & "Sub Window_OnLoad" & vbcrlf
        strSub = strSub & "    iTimerID = window.setInterval(" & chr(34) & "RunScript" & chr(34) & ", 5000, " & chr(34) & "VBScript" & chr(34) & ")" & vbcrlf
        strSub = strSub & "End Sub"        
        strExample = strScript
        strDescription = "Sample code that causes the script named RunScript to run every 5 seconds (5000 milliseconds). If you place any created timers in the Windows_OnLoad subroutine, those timers will automatically start when your HTA loads. As a result, any scripts they reference will automatically run at the specified intervals."
    End If

    If oFormHelpomatic.hta_elements.Value = "Table (2 Column)" Then
        oFormHelpomatic.reset_button.Disabled = True
        oFormHelpomatic.copy_code_button.Disabled = True
        strScript = "<table width=" & chr(34) & "100%" & chr(34) & " border>" & vbCrLf
        strScript = strScript & "  <tr>" & vbCrLf
        strScript = strScript & "    <td width=" & chr(34) & "50%" & chr(34) & " valign=" & chr(34) & "top" & chr(34) & ">Row 1, Column 1</td>" & vbCrLf
        strScript = strScript & "    <td width=" & chr(34) & "50%" & chr(34) & " valign=" & chr(34) & "top" & chr(34) & ">Row 1, Column 2</td>" & vbCrLf
        strScript = strScript & "  </tr>" & vbCrLf
        strScript = strScript & "</table>" & vbCrLf
        strExample = strScript
        strDescription = "Tables provide an easy way to create nicely-aligned, nicely-formatted output. This sample code creates a basic two-column table, with the two columns the same size. To create a table with different-sized columns, adjust the cell (td) widths accordingly. <br>&nbsp;<br>For tables (table), table rows (tr), and table cells (td), you can change the horizontal alignment (align), the vertical alignment (valign), the background color (bgcolor), the border color (bordercolor) and many other properties."
    End If

    If oFormHelpomatic.hta_elements.Value = "Table (3 Column)" Then'
        oFormHelpomatic.reset_button.Disabled = True
        oFormHelpomatic.copy_code_button.Disabled = True
        strScript = "<table width=" & chr(34) & "100%" & chr(34) & " border>" & vbCrLf
        strScript = strScript & "  <tr>" & vbCrLf
        strScript = strScript & "    <td width=" & chr(34) & "33%" & chr(34) & " valign=" & chr(34) & "top" & chr(34) & " border= " & chr(34) & "black" & chr(34) & ">Row 1, Column 1</td>" & vbCrLf
        strScript = strScript & "    <td width=" & chr(34) & "33%" & chr(34) & " valign=" & chr(34) & "top" & chr(34) & ">Row 1, Column 2</td>" & vbCrLf
        strScript = strScript & "    <td width=" & chr(34) & "33%" & chr(34) & " valign=" & chr(34) & "top" & chr(34) & ">Row 1, Column 3</td>" & vbCrLf
        strScript = strScript & "  </tr>" & vbCrLf
        strScript = strScript & "</table>" & vbCrLf
        strExample = strScript
        strDescription = "Tables provide an easy way to create nicely-aligned, nicely-formatted output. This sample code creates a basic three-column table, with each column the same size. To create a table with different-sized columns, adjust the cell (td) widths accordingly. <br>&nbsp;<br>For tables (table), table rows (tr), and table cells (td), you can change the horizontal alignment (align), the vertical alignment (valign), the background color (bgcolor), the border color (bordercolor) and many other properties."
    End If

    If oFormHelpomatic.hta_elements.value = "Table (4 Column)" Then
        oFormHelpomatic.reset_button.Disabled = True
        oFormHelpomatic.copy_code_button.Disabled = True
        strScript = "<table width=" & chr(34) & "100%" & chr(34) & " border>" & vbCrLf
        strScript = strScript & "  <tr>" & vbCrLf
        strScript = strScript & "    <td width=" & chr(34) & "25%" & chr(34) & " valign=" & chr(34) & "top" & chr(34) & ">Row 1, Column 1</td>" & vbCrLf
        strScript = strScript & "    <td width=" & chr(34) & "25%" & chr(34) & " valign=" & chr(34) & "top" & chr(34) & ">Row 1, Column 2</td>" & vbCrLf
        strScript = strScript & "    <td width=" & chr(34) & "25%" & chr(34) & " valign=" & chr(34) & "top" & chr(34) & ">Row 1, Column 3</td>" & vbCrLf
        strScript = strScript & "    <td width=" & chr(34) & "25%" & chr(34) & " valign=" & chr(34) & "top" & chr(34) & ">Row 1, Column 4</td>" & vbCrLf
        strScript = strScript & "  </tr>" & vbCrLf
        strScript = strScript & "</table>" & vbCrLf
        strExample = strScript
        strDescription = "Tables provide an easy way to create nicely-aligned, nicely-formatted output. This sample code creates a basic four-column table, with each column the same size. To create a table with different-sized columns, adjust the cell (td) widths accordingly. <br>&nbsp;<br>For tables (table), table rows (tr), and table cells (td), you can change the horizontal alignment (align), the vertical alignment (valign), the background color (bgcolor), the border color (bordercolor) and many other properties."
    End If

    If oFormHelpomatic.hta_elements.Value = "Table (Formatted Rows)" Then
        oFormHelpomatic.reset_button.Disabled = True
        oFormHelpomatic.copy_code_button.Disabled = True
        strScript = "<table width=" & chr(34) & "100%" & chr(34) & ">" & vbCrLf
        strScript = strScript & "<tr bgcolor=" & chr(34) & "white" & chr(34) & " align=" & chr(34) & "center" & chr(34) & " valign=" & chr(34) & "top" & chr(34) & ">"  & vbCrLf
        strScript = strScript & "    <td>Row 1, Column 1</td>" & vbCrLf
        strScript = strScript & "    <td>Row 1, Column 2</td>" & vbCrLf
        strScript = strScript & "  </tr>" & vbCrLf
        strScript = strScript & "  <tr bgcolor=" & chr(34) & "yellow" & chr(34) & " align=" & chr(34) & "right" & chr(34) & " valign=" & chr(34) & "top" & chr(34) & ">" & vbCrLf
        strScript = strScript & "    <td>Row 1, Column 1</td>" & vbCrLf
        strScript = strScript & "    <td>Row 1, Column 2</td>" & vbCrLf
        strScript = strScript & "  </tr>" & vbCrLf
        strScript = strScript & "</table>"    
        strExample = strScript
        strDescription = "Tables provide an easy way to create nicely-aligned, nicely-formatted output. This sample code creates a basic two-column, two-row table, with the two columns the same size, but with different alignments and background colors. To create a table with different-sized columns, adjust the cell (td) widths accordingly. <br>&nbsp;<br>For tables (table), table rows (tr), and table cells (td), you can change the horizontal alignment (align), the vertical alignment (valign), the background color (bgcolor), the border color (bordercolor) and many other properties. For example, add <b>rules = all </b> to the table tag and notice the difference."
    End If

    If oFormHelpomatic.hta_elements.Value = "Text Area" Then
        strScript = "<textarea name=" & chr(34) & "BasicTextArea" & chr(34) & " rows=" & chr(34) & "5" & chr(34) & " cols=" & chr(34) & "75" & chr(34) & "></textarea>"
        strSub = "'Places text in the text area named BasicTextArea" & vbcrlf & vbcrlf  
        strSub = strSub & "Sub RunScript" & vbcrlf
        strSub = strSub & "    BasicTextArea.Value = " & chr(34) & "This information will be placed in the multi-line text box named BasicTextArea." & chr(34) & vbcrlf
        strSub = strSub & "End Sub" 
        strExample = strScript & "&nbsp;&nbsp;&nbsp;<input id=runbutton  class=" & chr(34) & "button" & chr(34) & " type=" & chr(34) & "button" & chr(34) & " value=" & chr(34) & "Run Button" & chr(34) & " name=" & chr(34) & "run_button" & chr(34) & "  onClick=" & chr(34) & "RunScript" & chr(34) & ">"
        strDescription = "Textareas are simply multi-line text boxes; as a result, you can place a large amount of text information in a single text box. When configuring a textarea, you can specify the number of rows displayed on screen as well as the number of columns (one pixel per column). If the data being added exceeds the on-screen size of the textarea, scroll bars will automatically appear, allowing you to scroll up and down and view all the information.<br>&nbsp;<br>The sample script shows how you can programmatically add information to a textarea. You can also retrieve information from a textarea by retrieving the <b>Value</b>."
    End If

    If oFormHelpomatic.hta_elements.Value = "Text Box" Then
        strScript = "<input type=" & chr(34) & "text" & chr(34) & " name=" & chr(34) & "BasicTextBox" & chr(34) & " size=" & chr(34) & "50" & chr(34) & ">"
        strSub = "'Displays the contents of the text box BasicTextBox" & vbcrlf & vbcrlf  
        strSub = strSub & "Sub RunScript" & vbcrlf
        strSub = strSub & "    Msgbox BasicTextBox.Value" & vbcrlf
        strSub = strSub & "End Sub"
        strExample = strScript & "&nbsp;&nbsp;&nbsp;<input id=runbutton  class=" & chr(34) & "button" & chr(34) & " type=" & chr(34) & "button" & chr(34) & " value=" & chr(34) & "Run Button" & chr(34) & " name=" & chr(34) & "run_button" & chr(34) & "  onClick=" & chr(34) & "RunScript" & chr(34) & ">"
        strDescription = "Typically used to provide a way for users to enter information; users can type information (such as a user name) in the text box, and then you can programtically retrieve that information by checking the <b>Value</b> of the text box. A text box can be no bigger than a single line; to create larger text boxes, use the <b>Textarea</b> tag. To set a default for the text box, specify the Value as part of the input tag. To configure the size of the text box, set the <b>Width</b> as part of the input tag.<br>&nbsp;<br>The sample code simply echoes the information currently in the text box."
    End If

    If oFormHelpomatic.hta_elements.Value="ToolTip" Then
        strScript = "<input id=runbutton  class=" & chr(34) & "button" & chr(34) & " type=" & chr(34) & "button" & chr(34) & " value=" & chr(34) & "Run Button" & chr(34) & " name=" & chr(34) & "run_button" & chr(34) & "  onClick=" & chr(34) & "RunScript" & chr(34) & " title=" & chr(34) & "Click here to change the tooltip." & chr(34) & ">"
        strSub = "'Changes the tooltip for the button named run_button" & vbcrlf & vbcrlf  
        strSub = strSub & "Sub RunScript" & vbcrlf
        strSub = strSub & "    Run_Button.Title = " & chr(34) & "You successfully changed the tooltip." & chr(34) & vbcrlf
        strSub = strSub & "End Sub"
        strExample = strScript 
        strDescription = "Tooltips are little bits of explanatory information that appear whenever you hold the mouse above a control. Tooltips can be configured for most HTML elements simply by setting the <b>Title</b> property.<br>&nbsp;<br>In the example shown here, hold the mouse over the Run Button to see the current tooltip. Then, click the Run Button, and hold the mouse over it again; you should see that the tooltip has changed to the Title specified in the script code."
    End If

    oFormHelpomatic.ConfigureArea.Value = strScript
    oFormHelpomatic.SubroutineArea.Value = strSub
    SampleArea.InnerHTML = strExample
    oDivHelpDescription.InnerHTML = strDescription
    oScrExampleScript.Text = oFormHelpomatic.SubroutineArea.Value
    
End Sub

Sub ShowExample
    On Error Resume Next
    strScript = ""
    x = 0
    If oFormHelpomatic.hta_elements.Value="Checkbox" Then 
        x = 1
    End If
    If oFormHelpomatic.hta_elements.Value = "Disabled control" Then
        x = 1
    End If
    If oFormHelpomatic.hta_elements.Value = "Multi-select Listbox" Then
        x = 1
    End If
    If oFormHelpomatic.hta_elements.Value="Password box" Then
        x = 1
    End If
    If oFormHelpomatic.hta_elements.Value="Radio button" Then
        x = 1
    End If
    If oFormHelpomatic.hta_elements.Value = "Span" Then
        x = 1
    End If
    If oFormHelpomatic.hta_elements.Value = "Text Area" Then
        x = 1
    End If
    If oFormHelpomatic.hta_elements.Value = "Text Box" Then
        x = 1
    End If
    If x = 1 Then
        strScript = "<input id=runbutton  class='cInputBG' type='button' value='Run Button' name='run_button'  onClick='RunScript()'>"
    End If
    strScript = oFormHelpomatic.ConfigureArea.Value & "&nbsp;&nbsp;&nbsp;" & strScript
    SampleArea.InnerHTML = strScript
End Sub


Sub ResetCode
    On Error Resume Next
    oScrExampleScript.Text = oFormHelpomatic.SubroutineArea.Value
End Sub