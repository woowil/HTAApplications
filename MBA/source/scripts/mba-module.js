// nOsliw HUI - HTML/HTA Application Framework Library (https://github.com/woowil/HTAFrameworks)
// Copyright (c) 2003-2013, nOsliw Solutions,  All Rights Reserved.
//**Start Encode**

__H.register(__H.UI.Window.HTA.MBA,"_Module","Class for MBA modules",function _Module(){
	
	var bIsInitialized = false
	
	var add = function add(sFile,iIndex,sSubFolder){
		try{
			oXml.load(sFile)
			/*
			var oError = oXml.validate();
			if(oError.errorCode != 0){
				__HLog.logPopup("Unable to load XML file " + sFile + ". " + __HLog.trim(oError.reason))
				return;
			}*/
			var Dummy = oXml.documentElement.selectSingleNode("module_config")
		}
		catch(ee){
			__HLog.logPopup("_Module.add(): Unable to load XML file " + sFile)
			return;
		}

		this.module_file = sFile
		this.PathRoot = __HMBA.pth.mdl_name + "\\" + sSubFolder
		this.module_folder = sSubFolder

		var oNode = oXml.documentElement.selectSingleNode("module_revision")
		for(var oEnum = new Enumerator(oNode.attributes); !oEnum.atEnd(); oEnum.moveNext()){
			var oItem = oEnum.item();
			this[oItem.name] = oItem.nodeValue
		}

		var oNode = oXml.documentElement.selectSingleNode("module_information")
		for(var oEnum = new Enumerator(oNode.attributes); !oEnum.atEnd(); oEnum.moveNext()){
			var oItem = oEnum.item();
			this[oItem.name] = oItem.nodeValue
		}

		var idx = sSubFolder.indexOf("\\")
		
		this.menu_mm = "m_modules_" + (sSubFolder.substring(0,idx)).toLowerCase()
		this.menu_ii = "i_modules_" + (sSubFolder.substring(0,idx)).toLowerCase()

		this.menu_m = "m_" + this.shortname
		this.menu_i = "i_" + this.shortname

		var oElements = oXml.documentElement.getElementsByTagName("config")
		var oRe = /file|path/ig
		for(var i = oElements.length-1; i >= 0; i--){
			var t = oElements[i].getAttribute("type")
			if(t == "function") this[oElements[i].getAttribute("name")] = oElements[i].getAttribute("value")
			else if(t.isSearch(oRe)) this[oElements[i].getAttribute("name")] = this.PathRoot + "\\" + oElements[i].getAttribute("value")
			else this[oElements[i].getAttribute("name")] = oElements[i].getAttribute("value")
		}

		this.index = iIndex
	}

	this.initialize = function initialize(bLoadDefault){
		try{
			if(bIsInitialized){
				if(!confirm("Modules are already initialized (probably during startup).\nClick OK if you want to choose a new module path.")){
					return;
				}
			}
			
			var sFolder = __HMBA.pth.mdl
			if(!bLoadDefault){
				window.setTimeout(function(){mnu_hide_context()},0) // Prevent menu show while opening
				sFolder = __HShell.browseFolder(__HMBA.pth.mdl,"Path: " + __HMBA.pth.mdl)
				if(!sFolder) return;
			}
			
			__HSelect.setSelect(oFormModule.group)			
			__HSelect.addArray(__HFolder.list(__HMBA.pth.mdl,true),0)
			
			__HLog.logInfo("# Updating modules using folder: " + sFolder)
			
			var sModule = oFso.GetFileName(__HMBA.con_FileModule)			
			var aFiles = __HFile.list(sFolder,"xml",true,sModule)						
			if(aFiles.length == 0){
				__HLog.popup("Unable to find XML module files: " + sModule + " in folder: " + sFolder)
				return false
			}
			
			var oRe = new RegExp(__HMBA.pth.mdl.replace(/\\/g,"\\\\") + "\\\\(.+)\\\\" + sModule + "$","ig")
						
			for(var i = aFiles.length; i; i--){
				var o_div, o, t1, t2
				t1 = aFiles.shift()
				if(!(o = new add(t1,__HMBA.mdl.length,t1.replace(oRe,"$1")))) continue;
				else if(__HMBA.mdl_dic1.Exists(o.name)){
					__HLog.log("## Skipping module " + o.title + ". Already added!")
					continue;
				}
				
				o_div = __H.byClone("div")
				o_div.id = ("oDiv" + o.name + "_" + (10).random(99)).replace(/[ \t]*/g,"")
				oDivBody.appendChild(o_div)
				o.divname = o_div.id
				o.div = __H.byIds(o.divname)
				o.PathPicUrl = (o.PathPic).replace(/\\/g,"/")

				if(!__HMBA.mdl_dic3.Exists(o.menu_ii)){
					t1 = __H.byClone("optgroup")
					t2 = __H.byClone("optgroup")
					t1.label = t2.label = o.vendor
					oFormModule.load_list.appendChild(t1)
					oFormModule.unload_list.appendChild(t2)
					__HMBA.mdl_dic3.Add(o.menu_ii,t1)
					__HMBA.mdl_dic3.Add(o.menu_mm,t2)
				}
				t1 = __H.byClone("option")
				t1.innerHTML = o.title
				t1.value = "__MModule.load(" + o.index + ")"
				__HMBA.mdl_dic3(o.menu_ii).appendChild(t1)

				__HMBA.mdl_dic1.Add(o.name,o)
				__HMBA.mdl.push(o)
			}
		
			bIsInitialized = true
			return true
		}
		catch(ee){
			__HLog.error(ee,this)
			return false
		}
		finally{
			this.busy(false)
			__HMBA.isactive = false
			__HMBA.module_loaded = true
		}
	}

	this.load = function load(iIndex){
		try{
			if(!__HMBA.isready || __HMBA.isactive) return;
			var iMessage = 0
			this.busy(true)
			__HMBA.isactive = true
			
			if(!oFso.FileExists(__HMBA.mdl[iIndex].FileHTML)){
				__MModule.message(2)
				return false
			}
			
			__HMBA.mdl_cur = __HMBA.mdl[iIndex]

			var sFolderJS  = (__HMBA.mdl_cur.PathScript).replace(/^[.\\]*(.+)$/g,"$1")
			var sFolderCSS = (__HMBA.mdl_cur.PathStyle).replace(/^[.\\]*(.+)$/g,"$1")
			
			__HLog.logInfo("## Including HTML file for module: " + __HMBA.mdl_cur.title)
			var oFile = oFso.OpenTextFile(__HMBA.mdl_cur.FileHTML,__HIO.read,false,__HIO.TristateUseDefault);
			__HMBA.mdl_cur.div.style.display = "none"
			__HMBA.mdl_cur.div.innerHTML = oFile.ReadAll()			
			oFile.close()
			
			__HLog.logInfo("# Loading scripts and style for module: " + __HMBA.mdl_cur.title)
			if(__HLoad.loadScriptFolder(sFolderJS)){
				__HLoad.loadStyleFolder(sFolderCSS)
				
				__HLog.logInfo("## Creating menu for module: " + __HMBA.mdl_cur.title)
				if(!__HMBA.mnu[__HMBA.mdl_cur.menu_mm]){
					__HMBA.mnu["m_module"].mnu_add_item(new mnu_obj_menuItem(__HMBA.mdl_cur.vendor,__HMBA.mdl_cur.menu_ii,"code:"),"","explorer_16x16.gif");
					__HMBA.mnu[__HMBA.mdl_cur.menu_mm] = new mnu_obj_menu(170);
					__HMBA.mnu["m_module"].items[__HMBA.mdl_cur.menu_ii].mnu_set_sub(__HMBA.mnu[__HMBA.mdl_cur.menu_mm]);
				}
				__HMBA.mnu[__HMBA.mdl_cur.menu_mm].mnu_add_item(new mnu_obj_menuItem(__HMBA.mdl_cur.title,__HMBA.mdl_cur.menu_i,"code:"));

				window[__HMBA.mdl_cur.f_menubar]()
				window[__HMBA.mdl_cur.f_menubar] = __H.emptyFn

				window[__HMBA.mdl_cur.f_forminit]()
				window[__HMBA.mdl_cur.f_forminit] = __H.emptyFn
			
				for(var i = __HMBA.mdl_cur.div.childNodes.length; i; i--){
					var o = (__HMBA.mdl_cur.div).childNodes(i-1)
					if(o.tagName == "DIV") o.className = "mba-div-body"
					else iMessage = 1
					__HElem.hide(o) // Hide all unwanted tags
				}

				window[__HMBA.mdl_cur.f_init]()

				__HLog.logInfo("## Merging documentation for module: " + __HMBA.mdl_cur.title)
				if(oFso.FileExists(__HMBA.mdl_cur.FileDocs)){
					//__HFile.merge(__HMBA.fls.docs_ini,__HMBA.mdl_cur.FileDocs)
				}

				__HMBA.mdl_dic2.Add(__HMBA.mdl_cur.name,__HMBA.mdl_cur.index)
				window.setTimeout(function(){mba_common_formbehavior('behavior',__HMBA.mdl_cur.div,true)},0)
				//window.setTimeout(function(){mba_common_setcolor(__HMBA.color[0])},5)
		
				if(iMessage) __MModule.message(iMessage)

				var opt = oFormModule.load_list.options[oFormModule.load_list.selectedIndex]
				__HMBA.mdl_dic3(__HMBA.mdl_cur.menu_mm).appendChild(opt)

				oFormModule.unload.disabled = oFormModule.unloadall.disabled = false
				__HMBA.mdl_cur.div.style.display = "block"
			}
			else {
				__HLog.logPopup("## Scripts or Stylesheets failed loading")
			}

		}
		catch(ee){
			__HLog.error(ee,this);
		}
		finally{
			__HMBA.isactive = false
			this.busy(false)
		}
	}
	
	this.loadall = function loadall(){
		for(var j = __HMBA.mdl.length, i = j - 1; i >= 0; i--){
			if(!__HMBA.mdl_dic2.Exists(__HMBA.mdl[i].name)){
				this.load(i)
				//window.setTimeout("__MModule.load(i)",0)
				__HUtil.sleep(200)
			}
		}
		__HLog.popup((j-i) + "(" + j + ") modules were loaded successfully")
		window.setTimeout("__HMBA.refresh()",2000)
	}

	this.unload = function unload(iIndex){
		try{
			if(!__HMBA.isready) return;
			this.busy(true)
			
			var o = __HMBA.mdl[iIndex]
			__HLog.logInfo("# Unload module " + o.title)
			
			__HLoad.removeScripts(o.module_folder)
			mnu_obj_menudelete(__HMBA.mnu[o.menu_mm])
			o.div.removeNode(true)

			var opt = oFormModule.unload_list.options[oFormModule.unload_list.selectedIndex]
			__HMBA.mdl_dic3(__HMBA.mdl_cur.menu_ii).appendChild(opt)

			__HMBA.mnu["m_module"].delSubMenu(o.menu_ii)
			__HMBA.mdl_dic2.Remove(o.name)
		}
		catch(ee){
			__HLog.error(ee,this);
		}
		finally{
			this.busy(false)
		}
	}

	this.unloadall = function unloadall(){
		var oSel = oFormModule.unload_list
		for(var i = oSel.options.length; i; i--){
			oSel.selectedIndex = i-1
			oFormModule.unload.onclick()
			__HUtil.sleep(1000)
		}
	}
	
	this.create = function create(){
		try{
			
		}
		catch(e){
			__HLog.error(e,this);
			return false
		}
	}
	
	this.busy = function(bBusy){
		if(bBusy){
			oFormModule.unload_list.disabled = oFormModule.load_list.disabled = true
			oFormModule.unloadall.disabled = oFormModule.loadall.disabled = true
			oFormModule.unload.disabled = oFormModule.load.disabled = true
			oFormModule.create.disabled = true
			for(var i = oFormAction.action.length; i; i--) oFormAction.action[i-1].disabled = true
		}
		else {
			oFormModule.unload_list.disabled = oFormModule.load_list.disabled = false
			oFormModule.unloadall.disabled = oFormModule.loadall.disabled = false
			oFormModule.unload.disabled = oFormModule.load.disabled = false
			oFormModule.create.disabled = false
			for(var i = oFormAction.action.length; i; i--) oFormAction.action[i-1].disabled = false
		}
	}
	
	////////////////
	////// DEFAULT MODULE FUNCTIONS
	
	this.formactivate = function formactivate(oElem){
		try{
			if(!__HMBA.isready) return;
			for(var i = 0; i < __HMBA.mdl.length; i++){
				if(__HMBA.mdl_dic2.Exists(__HMBA.mdl[i].name)){
					window[__HMBA.mdl[i].f_formactivate](oElem)
				}
			}
			return true
		}
		catch(e){
			__HLog.error(e,this);
			return false
		}
	}

	this.formbehavior = function formbehavior(sOpt){
		try{
			if(!__HMBA.isready) return;
			for(var i = __HMBA.mdl.length; i; i--){
				if(__HMBA.mdl_dic2.Exists(__HMBA.mdl[i-1].name)){
					window[__HMBA.mdl[i-1].f_formbehavior](sOpt)
				}
			}
			return true
		}
		catch(e){
			__HLog.error(e,this);
			return false
		}
	}

	this.navigate = function navigate(sOpt,sOpt2,sOpt3){
		try{
			if(!__HMBA.isready) return;
			for(var i = __HMBA.mdl.length-1; i >= 0; i--){
				if(__HMBA.mdl_dic2.Exists(__HMBA.mdl[i].name)){
					if(window[__HMBA.mdl[i].f_navigate](sOpt,sOpt2,i)){
						__HMBA.mdl_cur = __HMBA.mdl[i]
						break;
					}
				}
			}
			return true
		}
		catch(e){
			__HLog.error(e,this);
			return false
		}
	}

	this.event = function event(oEvent){
		try{
			if(!__HMBA.isready) return;
			for(var i = __HMBA.mdl.length-1; i >= 0; i--){
				if(__HMBA.mdl_dic2.Exists(__HMBA.mdl[i].name)){
					window[__HMBA.mdl[i].f_event](oEvent)
				}
			}
			return true
		}
		catch(e){
			__HLog.error(e,this);
			return false
		}
	}

	this.message = function message(iOpt){
		try{
			var s = false
			switch(iOpt){
				case 1 : s = "Module import was successful. However, MBA hides HTML that were outside of main DIV tags"; break;
				case 2 : s = "Unable to locate HTML file from the XML module file. Please check the value 'FileHTML' in the module or config file."; break;
				default: break;
			}
			if(s) __HLog.logPopup(s)
		}
		catch(e){
			__HLog.error(e,this);
			return false
		}
	}

})

var __MModule = new __H.UI.Window.HTA.MBA._Module()