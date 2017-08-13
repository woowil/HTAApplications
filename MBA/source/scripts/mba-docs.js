// nOsliw HUI - HTML/HTA Application Framework Library (https://github.com/woowil/HTAFrameworks)
// Copyright (c) 2003-2013, nOsliw Solutions,  All Rights Reserved.
//**Start Encode**

__H.register(__H.UI.Window.HTA.MBA,"_Doc","Documents and dialogues",function _Doc(){	
	//  http://msdn2.microsoft.com/en-us/library/ms533723.aspx__H
	
	this.getFile = function(sFile){
		try{
			sFile = sFile ? sFile : __HIO.temp + "\\" + oFso.GetTempName()
			if(sFile = prompt("Save As",sFile)){
				return sFile
			}
			return false
		}
		catch(ee){
			__HLog.error(ee,this)
		}
	}
	
	this.dialogHelp = function dialogHelp(){
		try{
			showModalDialog(__HMBA.fls.docs_dialog_help,window,"dialogHeight:350px;dialogWidth:500px;edge:raised;help:no;resizable:no;application=yes;" );
		}
		catch(ee){
			__HLog.error(ee,this)
		}
	}
	
	this.dialogAbout = function dialogAbout(){
		try{
			if(typeof(mnu_hide_context) == "function") mnu_hide_context()
			showModalDialog(__HMBA.fls.docs_dialog_about,window,"dialogHeight:550px;dialogWidth:480px;edge:raised;help:no;resizable:no;application=yes;");
		}
		catch(ee){
			__HLog.error(ee,this)
		}
	}
	
	this.dialogBusy = function dialogBusy(){
		try{
			var iLeft = (window.screen.width-400)/2;
			var iTop = (window.screen.height-120)/2;
			showModalDialog(__HMBA.fls.docs_dialog_busy,window,"dialogLeft:" + iLeft + "px;dialogTop:" + iTop + "px;dialogHeight:120px;dialogWidth:400px;help:0;resizable:no;center:yes;edge:sunken;status:no;scroll:no;application=no;");	
		}
		catch(ee){
			__HLog.error(ee,this)
		}
	}
	
	this.dialogDebug = function dialogDebug(){
		try{
			if(typeof(mnu_hide_context) == "function") mnu_hide_context()
			showModalLessDialog(__HMBA.fls.docs_dialog_debug,window,"dialogHeight:550px;dialogWidth:480px;center:yes;status:yes;edge:raised;help:no;resizable:yes;application=yes;");
		}
		catch(ee){
			__HLog.error(ee,this)
		}
	}
	
	this.printFile = function printFile(sFile){
		try{
			oWsh.Run("notepad /p " + sFile,__HIO.hide,false);
		}
		catch(ee){
			__HLog.error(ee,this)
		}
	}
	
	this.printHTML = function printHTML(oElement){
		try{
			if(oElement.nodeName == "TEXTAREA") oElement.select();
			oElement.document.execCommand('print');
		}
		catch(ee){
			__HLog.error(ee,this)
		}
	}
	
	this.saveAsText = function saveAsText(oElement,sFile){
		try{
			if(!(sFile = this.getFile(sFile))) return;
			if(!__HFile.create(sFile,oElement.innerText));
			else {
				__HShell.open(sFile)
				return sFile;
			}
			return false
		}
		catch(ee){
			__HLog.error(ee,this)
		}
	}
	
	this.saveAsHTML = function saveAsHTML(oElement,sFile,iWidth,iHeight){
		try{
			if(!(sFile = this.getFile(sFile))) return;
			var sHtml = '<html><head><style type="text/css">@import url(' + __HMBA.fls.css_common_url + ')</style></head><body>\n' + oElement.innerHTML + '</body></html>'			
			sFile = __HFile.create(sFile,sHtml) ? sFile : "about:blank"
			delete sHtml
			
			iWidth = iWidth ? iWidth : 500
			iHeight = iHeight ? iHeight : 400
			
			var iLeft = (window.screen.width-iWidth)/2;
			var iTop = (window.screen.height-iHeight)/2;
			
			var sOptions = "height=" + iHeight + ",width=" + iWidth + ",left=" + iLeft + ",top=" + iTop + ",status=no,application=yes,titlebar=no,toolbar=no,menubar=no,location=no,scrollbars=no"
			var oWin = window.open(sFile,"saveAsHTML",sOptions)			
			oWin.moveTo(iLeft,iTop);
			oWin.resizeTo(iWidth,iHeight);
			
			//oWin.document.styleSheets(0).addImport(__HMBA.fls.css_common,0)
			oWin.focus()
		}
		catch(ee){
			__HLog.error(ee,this)
		}
	}
	
	this.getDocs = function getDocs(){
		try{
			if(!oFso.FileExists(__HMBA.fls.docs_ini)){
				__HLog.popup("Unable to locate information file: " + __HMBA.fls.docs_ini);
				return;
			}
			__HLog.log("# Getting documentation using INI file: " + __HMBA.fls.docs_ini)
			var oRow, oCell, aSections, aKeys
			var oTBody = __H.byIds("oBdyDocs1");
			if(!oTBody) return;
			for(var i = oTBody.rows.length; i > 0; i--) oTBody.deleteRow(i);
			
			__HINI.setFile(__HMBA.fls.docs_ini)
			if(aSections = __HINI.getSections()){
				for(var j = 0, i = aSections.length; i; i--, j++){
					__HINI.setSection(aSections.shift())
					if((aKeys = __HINI.getKeys()) && aKeys.length == 4){
						oRow = oTBody.insertRow(oTBody.rows.length), n = j.toNumberZero()
						oRow.style.verticalAlign = 'top';
						oCell = oRow.insertCell(), oCell.innerText = n, oCell.width = 22, oCell.align = "center"
						oCell = oRow.insertCell(), oCell.width = 110, oCell.className = "mba-table-1TD"
						oCell.innerText = aKeys[0].value + " "
						oCell = oRow.insertCell(), oCell.innerText = aKeys[1].value + " "
						oCell = oRow.insertCell(), oCell.innerText = aKeys[2].value + " "
						oCell = oRow.insertCell(), oCell.innerText = aKeys[3].value + " "
						oCell = oRow.insertCell(), oCell.innerText = n, oCell.width = 22, oCell.align = "center"
						aKeys.empty()
					}
				}
			}
			oTBody = __H.byIds("oBdyDocs2");
			oTBody.rows(0).style.verticalAlign = 'top';
			oTBody.rows(0).cells(0).innerHTML = "&nbsp;"
			oTBody.rows(0).cells(1).innerHTML = __HLang.disclaimer
			
			return true
		}
		catch(e){
			__HLog.error(e,this)
			return false;
		}
	}

})

var __MDoc = new __H.UI.Window.HTA.MBA._Doc()

function mba_obj_help(sSection,sPurpose,sRequirements,sComments){
	try{
		this.section = sSection
		this.purpose = sPurpose
		this.requirements = sRequirements
		this.comment = sComments
	}
	catch(e){
		__HLog.error(e,this)
		return false;
	}
}


