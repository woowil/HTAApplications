// Copyright© 2003-2008. nOsliw Solutions. All rights reserved.
//**Start Encode**

__H.include("HUI@Sys@Net_Lanman.js")

__H.register(__H.UI.Window.HTA.MBA._Module,"FileManager","Class for File Management",function FileManager(){
	var drives_ao = null
	var shares_ao = null
	var shell_folders_ao = null
	var isactive = false
	var isloaded = false
	var pic_path;
	var d_extention
	var s_current_folder = ""
	
	var oReJS = /js|jse|vbs|vbe|wsf|wsh/ig
	var oReZIP = /zip|cap|rar|tar|gz|7z/ig
	var oReEXE = /bat|cmd|exe|com|hta/ig
	var oReJPG = /jpeg|jpe|jpg|jfif/ig
	var oReTXT = /txt|log|asc|csv/ig
	var oReINI = /ini|inf/ig
	var oReHTM = /htm|html/ig

	var oReNotDirFile = /[\\/:\*\?<>|]/ig

	var oForm = null
	var bShowFolderSize = false
	var bShowAttributes = false
	var bShowInaccessable = false

	var ID_SHOWFILES

	var d_cells_marked = []	
	var current_index = null
	var d_last_cell = null

	//var d_selects = new ActiveXObject("Scripting.Dictionary")

	this.initialize = function initialize(){
		oForm = oFormManagementFile
		d_extention = new ActiveXObject("Scripting.Dictionary")
		d_cells_marked[0] = new ActiveXObject("Scripting.Dictionary")
		d_cells_marked[1] = new ActiveXObject("Scripting.Dictionary")
		isloaded = this.getDrives()
		return true
	}

	this.getDrives = function getDrives(){
		if(isactive) return;
		try{
			drives_ao = __HIO.getDrives()

			for(var i = 0; i < 2; i++){
				__HSelect.setSelect(oForm.list[i])
				__HSelect.addIndex(0,"mycomputer"," My Computer")
				__HSelect.addIndex(1,"myshared"," My Shared Folders")
				__HSelect.addIndex(2,"myspecial"," My Shell Folders")
				__HSelect.addArrayObject(drives_ao,true,3,"DriveLetter","VolumeName")
				__HSelect.setIndex(0)
			}

			return true
		}
		catch(ee){
			__HLog.error(ee,this)
			return false
		}
	}

	var setTBodyFirstRow = function setTBodyFirstRow(oTBody,bRoot,bMyOher,sFolder){
		try{
			for(var i = oTBody.rows.length-1; i >= 0; i--) oTBody.deleteRow(i);
			var oRow = oTBody.insertRow(oTBody.rows.length)
			var oCell = oRow.insertCell(), oFolder
			oCell.colSpan = 5, oCell.align = "left"
			oCell.style.color = __HMBA.color[__HMBA.cur_color_idx].col_dark
			oCell.innerHTML = "Directory of " + sFolder
			oCell.setAttribute("_isDirRoot",bRoot)
			oCell.setAttribute("_isDirMyOther",bMyOher)
			oCell.setAttribute("_DirPath",sFolder)

			if(oFso.FolderExists(sFolder) && (oFolder = oFso.GetFolder(sFolder))){
				oCell.setAttribute("_isDirSystem",(oFolder.attributes & 4))
				oCell.setAttribute("_isDirWritable",(bMyOher ? false : (oCell._isDirSystem && oFolder.IsRootFolder ? true : false)))
			}
		}
		catch(ee){
			__HLog.error(ee,this)
			return false
		}
	}

	this.showMyComputer = function showMyComputer(oTBody){
		try{
			oDivBody.style.cursor = "hand"
			setTBodyFirstRow(oTBody,false,true,"My Computer")

			drives_ao = __HIO.getDrives()
			for(var i = 0, iLen = drives_ao.length; i < iLen; i++){
				var oRow = oTBody.insertRow(oTBody.rows.length), oCell
				var o = drives_ao[i]
				var img = "<img width=16 height=16 src='" + this.getIconDrives(o) + "' border=0 align='absmiddle'> "
				oCell = oRow.insertCell(), oCell.innerHTML = img + "[" + o.DriveLetter + "] (" + o.VolumeName + ")"
				addEventCell(oCell,true), oCell.setAttribute("_DirPath",o.DriveLetter + "\\")
				oCell = oRow.insertCell(), oCell.innerText = o.AvailableSpace + " (" + o.TotalSize + ")"
				oCell = oRow.insertCell(), oCell.innerText = o.DriveTypeName + " (" + o.FileSystem + ")"
				oCell = oRow.insertCell(), oCell.innerText = " "
				oCell = oRow.insertCell(), oCell.innerText = o.DriveLetter
			}
			oTBody.parentNode.refresh()

			return true

		}
		catch(ee){
			__HLog.error(ee,this)
			return false
		}
		finally{
			oDivBody.style.cursor = "default"
		}
	}

	this.showMyShared = function showMyShared(oTBody){
		try{
			oDivBody.style.cursor = "hand"
			setTBodyFirstRow(oTBody,false,true,"My Shared Folders")

			var sQuery = "Select Description,Name,Path,Status from Win32_share where Status='OK'"
			var oColItems = (__HWMI.getServiceWMI(null,"root\\cimv2")).ExecQuery(sQuery,"WQL",48)

			var img = "<img width=16 height=16 src='" + __HMBA.mdl_cur.PathPicUrl + "/icon_Network Drive.ico' border=0 align='absmiddle'> "
			for(var oEnum = new Enumerator(oColItems); !oEnum.atEnd() && !__H.$stopit; oEnum.moveNext()){
				var oItem = oEnum.item();
				if(oItem.Name == "IPC$") continue
				var oRow = oTBody.insertRow(oTBody.rows.length), oCell
				var oFolder = oFso.GetFolder(oItem.Path)
				oCell = oRow.insertCell(), oCell.innerHTML = img + "[" + oItem.Name + "] (" + oItem.Path + ")"
				addEventCell(oCell,true), oCell.setAttribute("_DirPath",oItem.Path)
				setAttributesCell(oFolder,oCell)
				try{
					oCell = oRow.insertCell()
					oCell.innerText = bShowFolderSize ? __HSys.bytes(oFolder.Size) : " "
				}
				catch(ee){
					oCell.innerText = "[Access Denied]", oCell.style.color = "#FF0000"
				}
				oCell = oRow.insertCell(), oCell.innerText = oItem.Description
				oCell = oRow.insertCell(), oCell.innerText = oFolder.IsRootFolder ? " " : (new Date(oFolder.DateLastModified)).formatLocateString()
				oCell = oRow.insertCell(), oCell.innerText = oItem.Name
			}

			/*
			shares_ao = __HLanman.getShares(oWno.ComputerName)
			for(var i = 0; i < shares_ao.length; i++){
				var oRow = oTBody.insertRow(oTBody.rows.length), oCell
				var o = shares_ao[i]
				oCell = oRow.insertCell(), oCell.innerHTML = o.Name
				//addEventCell(oCell), oCell.setAttribute("_DirPath",o.DriveLetter + "\\")
				oCell = oRow.insertCell(), oCell.innerText = o.Class
				oCell = oRow.insertCell(), oCell.innerText = o.AdsPath
				oCell = oRow.insertCell(), oCell.innerText = o.HostComputer
				oCell = oRow.insertCell(), oCell.innerText = o.Path
			}
			*/

			oTBody.parentNode.refresh()

			return true

		}
		catch(ee){
			__HLog.error(ee,this)
			return false
		}
		finally{
			oDivBody.style.cursor = "default"
		}
	}

	this.showMyShell = function showMyShell(oTBody){
		try{
			oDivBody.style.cursor = "hand"
			setTBodyFirstRow(oTBody,false,true,"My Shell Folders")

			var sUserProfile = __HReg.getEnvironment("USERPROFILE","process")
			if(!(shell_folders_ao = __HReg.getValues(__HKeys.usr_shell))){
				throw new Error(8888,"Unable to get User Shell Folders")
			}
			var img = "<img width=16 height=16 src='" + __HMBA.mdl_cur.PathPicUrl + "/icon_dir.gif' border=0 align='absmiddle'> "
			for(var i = 0, iLen = shell_folders_ao.length; i < iLen; i++){
				var o = shell_folders_ao[i]				
				var sFolder = (o.Value).replace(/%userprofile%/ig,sUserProfile)
				try{
					oFolder = oFso.GetFolder(sFolder)
				}
				catch(ee){
					__HLog.logInfo("Unable to find Shell folder path: " + sFolder)
					continue
				}
				var oRow = oTBody.insertRow(oTBody.rows.length), oCell
				oCell = oRow.insertCell(), oCell.innerHTML = img + " " + o.Name
				addEventCell(oCell,true), oCell.setAttribute("_DirPath",sFolder)
				setAttributesCell(oFolder,oCell)
				try{
					oCell = oRow.insertCell()
					oCell.innerText = bShowFolderSize ? __HSys.bytes(oFolder.Size) : " "
				}
				catch(ee){
					oCell.innerText = "[Access Denied]", oCell.style.color = "#FF0000"
				}
				oCell = oRow.insertCell(), oCell.innerText = oFolder.Type
				oCell = oRow.insertCell(), oCell.innerText = (new Date(oFolder.DateLastModified)).formatLocateString()
				oCell = oRow.insertCell(), oCell.innerText = o.Value
			}

			oTBody.parentNode.refresh()

			return true
		}
		catch(ee){
			__HLog.error(ee,this)
			return false
		}
		finally{
			oDivBody.style.cursor = "default"
		}
	}

	this.showDirFiles = function showDirFiles(oTBody,sFolder){
		if(isactive && isloaded) return;
		try{
			isactive = true
			var iIndex = oTBody.index
			oDivBody.style.cursor = "hand"

			this.setForm(oTBody,sFolder)
			this.resetCells(oTBody,true)

			if(sFolder == "mycomputer") return this.showMyComputer(oTBody)
			else if(sFolder == "myshared") return this.showMyShared(oTBody)
			else if(sFolder == "myspecial") return this.showMyShell(oTBody)

			var oFolder = oFso.GetFolder(s_current_folder = oFso.GetAbsolutePathName(sFolder))
			
			if(oFolder.IsRootFolder){
				if(!oFolder.Drive.IsReady) return;
				if(oFolder != oFolder.Drive.RootFolder){
					// This forces current folder in any drive to point to root folder
					oFolder = oFolder.Drive.RootFolder;
				}
			}
			setTBodyFirstRow(oTBody,oFolder.IsRootFolder,false,s_current_folder)

			//if(!oFolder.IsRootFolder){
			var oRow = oTBody.insertRow(oTBody.rows.length), oCell
			var oParentFolder = oFso.GetFolder((oFolder.Path).replace(/(.+)\\.+/g,"$1"))

			oCell = oRow.insertCell()
			oCell.innerHTML = "<img width=16 height=16 src='" + __HMBA.mdl_cur.PathPicUrl + "/icon_dir_up.gif' border=0 align='absmiddle'> [..]"
			addEventCell(oCell,true)
			oCell.setAttribute("_DirPath",(!oFolder.IsRootFolder ? oParentFolder.Path : "mycomputer"))

			oCell = oRow.insertCell(), oCell.innerText = " "
			oCell = oRow.insertCell(), oCell.innerText = (!oFolder.IsRootFolder ? oFolder.Type : "My Computer")
			oCell = oRow.insertCell(), oCell.noWrap = true
			oCell.innerText = !oParentFolder.IsRootFolder ? (new Date(oParentFolder.DateLastModified)).formatLocateString() : " "
			oCell = oRow.insertCell()
			oCell.innerText = !oParentFolder.IsRootFolder ? oParentFolder.ShortName : " "
			//}
			// Listing of folders

			for(var oEnum = new Enumerator(oFolder.SubFolders); !oEnum.atEnd(); oEnum.moveNext()){
				var oItem = oEnum.item(), oCell
				var oRow = oTBody.insertRow(oTBody.rows.length)

				oCell = oRow.insertCell()
				oCell.setAttribute("_DirPath",oItem.Path)
				oCell.innerHTML = "<img width=16 height=16 src='" + __HMBA.mdl_cur.PathPicUrl + "/icon_dir.gif' border=0 align='absmiddle' alt='Double click to open'> <input type='test' name='" + iIndex + "_c_dir' value='" + oItem.Name + "' class='mba-input-tw'>"
				setAttributesCell(oItem,oCell), addEventCell(oCell)

				oCell = oRow.insertCell()
				try{
					oCell.innerText = bShowFolderSize ? __HSys.bytes(oItem.Size) : " "
				}
				catch(ee){
					oCell.innerText = "[Access Denied]", oCell.style.color = "#FF0000"
				}
				oCell = oRow.insertCell(), oCell.innerText = oItem.Type
				oCell = oRow.insertCell(), oCell.noWrap = true, oCell.innerText = (new Date(oItem.DateLastModified)).formatLocateString()
				oCell = oRow.insertCell(), oCell.innerText = oItem.ShortName
			}
			// Listing of files
			for(var oEnum  = new Enumerator(oFolder.Files); !oEnum.atEnd(); oEnum.moveNext()){
				var oItem = oEnum.item(), oCell
				var oRow = oTBody.insertRow(oTBody.rows.length)
				//__HLog.debug("## break2 " + oItem.Name)
				oCell = oRow.insertCell()
				oCell.setAttribute("_FilePath",oItem.Path)
				oCell.innerHTML = "<img width=16 height=16 src='" + this.getIconFiles(oItem.Name) + "' border=0 align='absmiddle' alt='Double click to open'> <input type='test' name='" + iIndex + "_c_file' value='" + oItem.Name + "' class='mba-input-tw'>"
				setAttributesCell(oItem,oCell), addEventCell(oCell)

				oCell = oRow.insertCell(), oCell.noWrap = true, oCell.innerText = __HSys.bytes(oItem.Size)
				oCell = oRow.insertCell(), oCell.innerText = oItem.Type
				oCell = oRow.insertCell(), oCell.noWrap = true, oCell.innerText = (new Date(oItem.DateLastModified)).formatLocateString()
				oCell = oRow.insertCell()
				try{
					// Bug!! Files like pagefile.sys or hyberfile.sys gives: Invalid procedure call or argument
					oCell.innerText = oItem.ShortName
				}
				catch(ee){
					oCell.innerText = (oFso.GetBaseName(oItem.Name)).length <= 8 ? oItem.Name : "N/A"
				}
			}

			oTBody.parentNode.refresh()
			return true
		}
		catch(ee){
			__HLog.error(ee,this)
			return false
		}
		finally{
			oDivBody.style.cursor = "default"
			isactive = false
		}
	}

	var setAttributesCell = function setAttributesCell(oItem,oCell){
		if(oItem.attributes & 2048){
			oCell.style.color = "blue"
			oCell.setAttribute("_isCompress",true)
		}
		if(oItem.attributes & 4){
			//oCell.style.fontStyle = "italic"
			try{
				oCell.childNodes(2).style.color = "orange"
				oCell.childNodes(2).onselect = oCell.childNodes(2).onselectstart = new Function("return false")
				oCell.childNodes(2).onfocus = new Function("return false")
				oCell.childNodes(2).readonly = true
			}
			catch(ee){}
			oCell.setAttribute("_isSystem",true)
		}
		if(bShowAttributes){
			var s = new __H.Common.StringBuffer()
			if(oItem.attributes & 1) s.append("R")
			if(oItem.attributes & 2) s.append("H")
			if(oItem.attributes & 4) s.append("S")
			if(oItem.attributes & 32) s.append("A")
			oCell.innerHTML = (oCell.innerHTML).concat(" [" + s + "]")
		}
		if(bShowInaccessable){
			if(oCell._DirPath){
				oCell.setAttribute("_isReadable",(__HFolder.isReadable(oCell._DirPath) ? true : false))
			}
			else {
				oCell.setAttribute("_isReadable",(__HFile.isReadable(oCell._FilePath) ? true : false))
			}
			if(oCell._isReadable){
				oCell.style.color = "red"
			}
		}
	}

	var addEventCell = function addEventCell(oCell,bMyOther){
		oCell.onmouseover = new Function("__HFileMan.setEventCell('onmouseover',this)")
		oCell.onmouseout = new Function("__HFileMan.setEventCell('onmouseout',this)")
		if(!bMyOther) oCell.childNodes(0).ondblclick = new Function("__HFileMan.setEventCell('ondblclick',this.parentNode)") // the IMG
		else oCell.ondblclick = new Function("__HFileMan.setEventCell('ondblclick',this)") // the IMG
		oCell.onmousedown = new Function("__HFileMan.setEventCell('onmousedown',this)")
		oCell.onclick = new Function("__HFileMan.setEventCell('onclick',this)")
		if(!bMyOther) oCell.childNodes(2).onchange = new Function("__HFileMan.rename(this.parentNode)") // The Input text
	}

	this.resetCells = function resetCells(oTBody,bForce){
		try{
			if(!isloaded) return;
			if(bForce) current_index = oTBody.index
			if(bForce || current_index != null && oTBody.index != current_index){
				__HLog.debug("## Removing all marked cells in tbody index: " + current_index)
				var a = (new VBArray(d_cells_marked[current_index].Keys())).toArray()
				for(var i in a){
					if(!a.hasOwnProperty(i)) continue
					__HLog.debug("### Removing cell name: " + a[i])
					var oCell = d_cells_marked[current_index](a[i])
					this.resetCell(oCell,current_index,a[i])
					delete a[i]
				}
				this.setForm(oTBody)
			}
		}
		catch(ee){
			__HLog.error(ee,this)
		}
	}

	this.resetCell = function resetCell(oCell,iIndex){
		try{
			oCell.setAttribute("_isLock",false)
			oCell.onmouseout()
			if(d_cells_marked[iIndex].Exists(oCell._Name)){
				d_cells_marked[iIndex].Remove(oCell._Name)
			}
			oCell.childNodes(2).onselectstart = new Function("return true")
			oCell.childNodes(2).onselect = new Function("return true")
			oCell.childNodes(2).readonly = false
		}
		catch(ee){
			__HLog.error(ee,this)
		}
	}

	this.getLockCells = function getLockCells(oTBody){
		try{
			var a = []
			__HLog.debug("### Getting locked cells using index: " + oTBody.index)
			var a = (new VBArray(d_cells_marked[oTBody.index].Keys())).toArray()
			for(var i in a){
				if(!a.hasOwnProperty(i)) continue
				a.push(d_cells_marked[oTBody.index](a[i]))
				delete a[i]
			}
		}
		catch(ee){
			__HLog.error(ee,this)
		}
		return a
	}

	this.getLockDirFiles = function getLockDirFiles(oTBody){
		try{
			__HLog.debug("### Getting locked files using index: " + oTBody.index)
			var a = []
			var aCells = this.getLockCells(oTBody)
			for(var i = 0, iLen = aCells.length; i < iLen; i++){
				var sDirFile = aCells[i]._DirPath ? aCells[i]._DirPath : aCells[i]._FilePath
				a.push(sDirFile)
			}
		}
		catch(ee){
			__HLog.error(ee,this)
		}
		return a
	}

	this.setEventCell = function setEventCell(sOpt,oCell){
		if(isactive) return;
		try{
			var oTBody = oCell.parentNode.parentNode
			var iIndex = parseInt(oTBody.index)
			var iIndex2 = iIndex == 0 ? 1 : 0
			var oTBody2 = oTbyManagementFile[iIndex2]

			if(sOpt == "onmouseover"){
				if(oCell._isLock) return;
				var o = __HMBA.color[__HMBA.cur_color_idx]
				oCell.style.backgroundColor = o.col_normal
				oCell.style.color = o.col_dark
				oCell.style.borderColor = o.col_dark
				oCell.style.cursor = "hand"
				if(oCell.firstChild.src && (oCell.firstChild.src).isSearch(/icon_dir\.gif/ig)){
					oCell.firstChild.src = __HMBA.mdl_cur.PathPicUrl + "/icon_dir_open.gif"
				}
			}
			else if(sOpt == "onmouseout"){
				if(oCell._isLock) return;
				oCell.style.backgroundColor = '#EFEFEF'
				oCell.style.color = oCell._isSystem ? 'orange' : (oCell._isCompress ? '0000FF' : '#000000')
				oCell.style.cursor = "default"
				oCell.style.borderColor = "#EFEFEF"
				if(oCell.firstChild.src && (oCell.firstChild.src).isSearch(/icon_dir_open\.gif/ig)){
					oCell.firstChild.src = __HMBA.mdl_cur.PathPicUrl + "/icon_dir.gif"
				}
			}
			else if(sOpt == "ondblclick"){
				if(oCell._DirPath){
					return this.showDirFiles(oTBody,oCell.getAttribute("_DirPath"))
					//if(ID_SHOWFILES) clearTimeout(ID_SHOWFILES), ID_SHOWFILES = null
					//ID_SHOWFILES = setTimeout("__HFileMan.showDirFiles(oTBody,oCell.getAttribute("_DirPath")",100))
				}
				else if(oCell._FilePath){
					oWsh.Run("%comspec% /c start \"Must have title\" \"" + oCell._FilePath + "\"",__HIO.hide,false)
				}
			}
			else if(sOpt == "onclick"){
				var e = window.event
				if(isactive || !e.ctrlKey || oCell._isSystem || oTBody.rows(0).cells(0)._isDirMyOther) return;

				__HLog.debug("# Checking Tbody for index: " + iIndex)
				this.resetCells(oTBody)

				if(oCell._isLock){
					__HLog.debug("## Unlocking/Removing cell name: " + sName)
					this.resetCell(oCell,iIndex)
				}
				else {
					var sName = iIndex + "_" + (oCell._DirPath ? oCell._DirPath : oCell._FilePath).replace(/.+\\(.+)$/g,"$1")
					__HLog.debug("## Locking/adding cell name: " + sName)
					if(!d_cells_marked[iIndex].Exists(sName)){
						d_cells_marked[iIndex].Add(sName,oCell)
						d_last_cell = oCell
					}
					current_index = iIndex
					oCell.setAttribute("_isLock",true)
					oCell.setAttribute("_Name",sName)
					oCell.childNodes(2).onselectstart = new Function("return false")
					oCell.childNodes(2).onselect = new Function("return false")
					oCell.childNodes(2).readonly = true
				}
				this.setAction(oTBody,oCell)
			}
			else if(sOpt == "onmousedown"){

				event.returnValue = true;
			}
		}
		catch(ee){
			__HLog.error(ee,this)
			return false
		}
	}

	this.getIconDrives = function getIconDrives(oDrive){
		try{
			sDrvType = oDrive.DriveLetter == "A:" ? "Floppy Drive" : oDrive.DriveTypeName
			if(!d_extention.Exists(sDrvType)){
				var sFileName = "icon_" + sDrvType + ".ico"
				if(!oFso.FileExists(__HMBA.mdl_cur.PathPicUrl + "/" + sFileName)){
					sFileName = "icon_none.gif"
				}
				d_extention.Add(sDrvType,__HMBA.mdl_cur.PathPicUrl + "/" + sFileName)
			}
			return d_extention(sDrvType)
		}
		catch(ee){
			__HLog.error(ee,this)
			return false
		}
	}

	this.getIconFiles = function getIconFiles(sFileName){
		try{
			var sExt = oFso.GetExtensionName(sFileName)
			if(!d_extention.Exists(sExt)){
				var sFileName = "icon_" + sExt + ".gif"
				if(!oFso.FileExists(__HMBA.mdl_cur.PathPicUrl + "//" + sFileName)){
					if(sExt.isSearch(oReJS)) sFileName = "icon_js.gif"
					else if(sExt.isSearch(oReZIP)) sFileName = "icon_zip.gif"
					else if(sExt.isSearch(oReEXE)) sFileName = "icon_bat.gif"
					else if(sExt.isSearch(oReJPG)) sFileName = "icon_jpg.gif"
					else if(sExt.isSearch(oReJPG)) sFileName = "icon_txt.gif"
					else if(sExt.isSearch(oReINI)) sFileName = "icon_ini.ico"
					else if(sExt.isSearch(oReHTM)) sFileName = "icon_htm.gif"
					else if(sExt.isSearch(/lnk/ig)) sFileName = "icon_lnk.ico"
					else sFileName = "icon_none.gif"
				}
				d_extention.Add(sExt,__HMBA.mdl_cur.PathPicUrl + "//" + sFileName)
			}
			return d_extention(sExt)
		}
		catch(ee){
			__HLog.error(ee,this)
			return false
		}
	}


	this.setForm = function setForm(oTBody,sFolder){
		try{
			var iIndex = oTBody.index

			bShowFolderSize = oForm.show_foldersize.checked
			bShowAttributes = oForm.show_attributes.checked
			bShowInaccessable = oForm.show_inaccessable.checked

			if(sFolder) oForm.list[iIndex].value = sFolder.replace(/^([a-z]:)\\.+$/ig,"$1")

			oForm.copy[iIndex].disabled = oForm.move[iIndex].disabled = oForm.remove[iIndex].disabled = true
			if(sFolder) oForm.makedir[iIndex].disabled = oForm.makefile[iIndex].disabled = sFolder.isSearch(/mycomputer|myshared|myshell/ig)

			return true
		}
		catch(ee){
			__HLog.error(ee,this)
			return false
		}
	}

	this.setAction = function setAction(oTBody,oCell){
		try{
			var iIndex = oTBody.index
			oForm.copy[iIndex].disabled = oForm.move[iIndex].disabled = oForm.remove[iIndex].disabled = oCell._isLock ? false : true
			oForm.makedir[iIndex].disabled = oForm.makefile[iIndex].disabled = oTBody.rows(0).cells(0)._isDirWritable ? false : true
		}
		catch(ee){
			__HLog.error(ee,this)
		}
	}

	this.reset = function reset(){
		try{

		}
		catch(ee){
			__HLog.error(ee,this)
		}
	}

	this.copy = function copy(obj){
		try{

		}
		catch(ee){
			__HLog.error(ee,this)
		}
	}

	this.move = function move(obj){
		try{

		}
		catch(ee){
			__HLog.error(ee,this)
		}
	}

	this.rename = function rename(oCell){
		try{
			var oInput = oCell.childNodes(2)
			var sName = oCell._DirPath ? (oCell._DirPath).replace(/.+\\(.+)$/g,"$1") : (oCell._FilePath).replace(/.+\\(.+)$/g,"$1")
			var bRenamed = false
			if((oInput.value).isSearch(oReNotDirFile) || (oInput.value).isEmpty()){
				alert("A file or directory cannot be empty or contain charactors: \\/:*?\"<>|")
				oInput.value = sName
			}
			else if(oFso.FileExists(oInput.value) || oFso.FolderExists(oInput.value)){
				alert("A file or directory with same name already exists!")
				oInput.value = sName
			}
			else if(oCell._DirPath){
				if(confirm("You have renamed the directory: " + oCell._DirPath + "\nOld\t" + sName + "\nNew\t" + oInput.value + "\n\nCommit changes?")){
					var sFolder = (oCell._DirPath).replace(/(.+)\\(.+)$/g,"$1\\" + oInput.value)
					bRenamed = __HFolder.move(oCell._DirPath,sFolder)
				}
			}
			else {
				if(confirm("You have renamed the file: " + oCell._FilePath + "\nOld\t" + sName + "\nNew\t" + oInput.value + "\n\nCommit changes?")){
					var sFile = (oCell._FilePath).replace(/(.+)\\(.+)$/g,"$1\\" + oInput.value)
					bRenamed = __HFile.move(oCell._FilePath,sFile)
				}
			}
			if(bRenamed) this.refresh(oCell.parentNode.parentNode)
			return bRenamed
		}
		catch(ee){
			__HLog.error(ee,this)
			return false
		}
	}

	this.makedir = function makedir(oTBody){
		try{
			var oRe = /[\\/:\*\?<>|]/ig, sName
			var sFolder = oTBody.rows(0).cells(0)._DirPath
			if(sName = prompt("Enter a new folder name in folder: " + sFolder,"New Directory")){
				if(sName.isSearch(oRe) || sName.isEmpty()){
					if(confirm("A folder cannot be empty or contain charactors: \\/:*?\"<>|\n\nTry again?")){
						return this.makedir(oTBody)
					}
				}
				else if(oFso.FolderExists(sFolder = sFolder + "\\" + sName)){
					if(confirm("The file " + sFolder + " already exists\n\nTry again?")){
						return this.makedir(oTBody)
					}
				}
				else {
					if(__HFolder.create(sFolder)){
						this.refresh(oTBody,sFolder)
					}
					else {
						__HLog.popup("Unable to create folder: " + sFolder)
					}
				}
				return oFso.FolderExists(sFolder)
			}
			return true
		}
		catch(ee){
			__HLog.error(ee,this)
			return false
		}
	}

	this.makefile = function makefile(oTBody){
		try{
			var oRe = /[\\/:\*\?<>|]/ig, sFile
			var sFolder = oTBody.rows(0).cells(0)._DirPath
			if(sFile = prompt("Enter a new file name in folder: " + sFolder,"New Text Document.txt")){
				if(sFile.isSearch(oRe) || sFile.isEmpty()){
					if(confirm("A file cannot be empty or contain charactors: \\/:*?\"<>|\n\nTry again?")){
						return this.makefile(oTBody)
					}
				}
				else if(oFso.FileExists(sFile = sFolder + "\\" + sFile)){
					if(confirm("The file " + sFile + " already exists\n\nTry again?")){
						return this.makefile(oTBody)
					}
				}
				else {
					if(__HFile.create(sFile)){
						oWsh.Run("%comspec% /c start \"Must have title\" \"" + sFile + "\"",__HIO.hide,false)
						this.refresh(oTBody,sFolder)
					}
					else {
						__HLog.popup("Unable to create file: " + sFile)
					}
				}
				return oFso.FileExists(sFile)
			}
			return true
		}
		catch(ee){
			__HLog.error(ee,this)
			return false
		}
	}

	this.remove = function remove(oTBody){
		if(oForm.remove[oTBody.index].disabled) return;
		try{
			var aDirFiles = this.getLockDirFiles(oTBody)
			var bRemoved = true

			if(aDirFiles.length == 0);
			else if(confirm("You have selected " + aDirFiles.length + " DirFiles for deletion.\nDo you wish to proceed?")){
				for(var i = 0, iLen = aDirFiles.length; i < iLen; i++){
					__HLog.debug("Deleting DirFile: " + aDirFiles[i])
					if(!__HFile.remove(aDirFiles[i]) && !__HFolder.remove(aDirFiles[i])){
						__HLog.log("## ERROR Unable to remove DirFile: " + aDirFiles[i])
						bRemoved = false
					}
				}
				this.refresh(oTBody,s_current_folder)
			}

			return bRemoved
		}
		catch(ee){
			__HLog.error(ee,this)
			return false
		}
	}

	this.refresh = function refresh(oTBody,sFolder){
		var sFolder = typeof(sFolder) == "string" ? sFolder : (oTBody.rows(0).cells(0)._DirPath).replace(/\\/g,"\\\\")
		__HFileMan.showDirFiles(oTbyManagementFile[oTBody.index],sFolder)
	}
})

var __HFileMan = new __H.UI.Window.HTA.MBA._Module.FileManager()