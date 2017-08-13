// Copyright© 2003-2008. nOsliw Solutions. All rights reserved.
//**Start Encode**

function management_init(){
	try{

		return true
	}
	catch(e){
		__HLog.error(e,this);
		return false
	}
}

function management_menu(){
	try{
		__HMBA.mnu[__HMBA.mdl_cur.menu_m] = new mnu_obj_menu(160);
		__HMBA.mnu[__HMBA.mdl_cur.menu_mm].items[__HMBA.mdl_cur.menu_i].mnu_set_sub(__HMBA.mnu[__HMBA.mdl_cur.menu_m]);
		with(__HMBA.mnu[__HMBA.mdl_cur.menu_m]){
			mnu_add_item(new mnu_obj_menuItem("Process Monitor","","code:mba_common_navigate('nav_management_process','Process Monitor')"),"");
			mnu_add_item(new mnu_obj_menuItem("WMI Management","","code:mba_common_navigate('nav_management_wmi','WMI Management')"),"");
			mnu_add_item(new mnu_obj_menuItem("File Management","","code:mba_common_navigate('nav_management_file','File Management')"),"");
		}


	}
	catch(e){
		__HLog.error(e,this)
	}
}

function management_main(){
	try{
		if(!management_formvalidate()) return false

		if(__HMBA.mod.cur_form.name == "oFormManagementProcess"){
			if(management_process()){
				__HElem.showImg(oDivManagementProcess1Sub,oDivManagementProcess1Sup.firstChild,__HMBA.pth.pic_url + "/body/black_up_16x16.png",true);
			}
			return true
		}
		else if(__HMBA.mod.cur_form.name == "oFormManagementWmi"){
			__HPanel.deactivate("oDivManagementWmiPanel",3);
			if(management_wmi()){
				__HPanel.activate("oDivManagementWmiPanel",3);
				return true
			}
		}
		else if(__HMBA.mod.cur_form.name == "oFormManagementFile"){
			return __HFileMan.inititialize()
		}
		return false
	}
	catch(e){
		__HLog.error(e,this);
		return false
	}
}

function management_forminit(){
	try{
		with(oFormManagementProcess){
			process_interval_stop.onclick = new Function("management_formactivate(this)");
			process_connect.onclick = new Function("management_formactivate(this)");
			process_filter_action.onclick = new Function("management_formactivate(this)");
			onSubmit = new Function("return false");
		}

		with(oFormManagementWmi){
			__HSelect.setSearch(oFormManagementWmi.wmi_namespace,0,false,'[0-9]{3} ')
			__HSelect.setSearch(oFormManagementWmi.wmi_classes,0,false,'[0-9]{3} ')

			wmi_scr_copy.onclick = new Function("__HElem.select(oFormManagementWmi.wmi_scr_code)")
			wmi_ope_action.onclick = new Function("management_formactivate(this);");
			wmi_namespace.onchange = new Function("management_formactivate(this);");
			wmi_classes.onchange = new Function("management_formactivate(this);");
			wmi_scr_language.onchange = new Function("management_formactivate(this);");
			wmi_scr_run.onclick = new Function("management_formactivate(this)")

			wmi_clear_met.onclick = new Function("management_cleartbody('oBdyManagementWmiMet')")
			wmi_clear_pro.onclick = new Function("management_cleartbody('oBdyManagementWmiPro')")
			wmi_saveas_pro.onclick = new Function("__MDoc.saveAsHTML(oDivManagementWmiPanel.childNodes(1).childNodes(3))")
			wmi_saveas_met.onclick = new Function("__MDoc.saveAsHTML(oDivManagementWmiPanel.childNodes(1).childNodes(2))")			
			wmi_print_met.onclick = new Function("__MDoc.printHTML(oDivManagementWmiPanel.childNodes(1).childNodes(2))")
			wmi_print_pro.onclick = new Function("__MDoc.printHTML(oDivManagementWmiPanel.childNodes(1).childNodes(3))")

			username.value = __HMBA.userdomain
			clear.onclick = new Function("this.form.reset();this.form.username.value=__HMBA.userdomain;this.form.server.value=__HMBA.comp")
			onsubmit = new Function("return false");
		}

		with(oFormManagementFile){

		}

		return true
	}
	catch(e){
		__HLog.error(e,this);
		return false
	}
}

function management_formbehavior(sOpt){
	try{
		if(sOpt == "management_process_reset"){
			var oForm = oFormManagementProcess;
			oForm.reset()
			oForm.server.value = oWno.ComputerName
			oForm.username.value = oWno.UserDomain + "\\" + oWno.UserName
			oForm.process_interval_stop.checked = false
			oForm.process_interval.value = "5000"
			oForm.process_classes.selectedIndex = 0
			var oCap = __H.byIds("oCapManagementProcess")
			oCap.innerHTML = "<i>System Process Monitor on</i>"
			setTimeout("management_cleartbody('oBdyManagementProcess')",0)
			__HElem.hideImg(oDivManagementProcess1Sub,oDivManagementProcess1Sup.firstChild,__HMBA.pth.pic_url + "/body/black_down_24x24.png",true);
		}
		else if(sOpt == "management_wmi_reset"){
			var oForm = oFormManagementWmi;
			oForm.reset()
			oForm.server.value = oWno.ComputerName
			oForm.username.value = oWno.UserDomain + "\\" + oWno.UserName
			oForm.query.value = document.all["oTextareaWmiScr"].innerText = ""
			oForm.wmi_opt_system.checked = oForm.wmi_opt_system.checked = false
			oForm.wmi_opt_win32.checked = true
			setTimeout("management_cleartbody('oBdyManagementWmiPro')",0)
			var oCap = __H.byIds("oCapManagementWmiPro")
			oCap.innerHTML = "<i>Properties for</i>"
			setTimeout("management_cleartbody('oBdyManagementWmiMet')",0)
			oCap = __H.byIds("oCapManagementWmiMet")
			oCap.innerHTML = "<i>Methods for</i>"
			oForm.wmi_ope_action.onclick();
			if(oFormAction.action) oFormAction.action[0].disabled = oFormAction.action[1].disabled = true
			__HPanel.deactivate("oDivManagementWmiPanel")
		}
		return true
	}
	catch(e){
		__HLog.error(e,this);
		return false
	}
}

function management_formactivate(oElement){
	try{
		var oForm = oElement.form
		if(oElement.form.name == "oFormManagementProcess"){
			if(oElement.name == "process_connect"){
				if(__H.isStringEmpty(oElement.form.server.value)) oElement.form.server.value = oWno.ComputerName
				__HMBA.wmi.u_server = oElement.form.server.value

				if(__H.isStringEmpty(oElement.form.username.value,oElement.form.password.value)) __HMBA.wmi.u_user = __HMBA.wmi.u_pass = null // Local connections can't use password in WMI!
				else __HMBA.wmi.u_user = oElement.form.username.value, __HMBA.wmi.u_pass = oElement.form.password.value

				__HMBA.wmi.u_namespace = "root\\CIMV2"

				__HLog.log("# Making a WMI Process Monitor connection to system " + __HMBA.wmi.u_server)
				if(!__HWMI.setServiceWMI(__HMBA.wmi.u_server,__HMBA.wmi.u_namespace,__HMBA.wmi.u_user,__HMBA.wmi.u_pass)) return false;
				
				setTimeout("oFormManagementProcess.process_filter_action.onclick()",0)
				return true
			}
			else if(oElement.name == "process_filter_action"){
				if(!__HWMI.getServiceWMI()) return;
				var a = __HWMI.getClassesSpecific(oElement.form.process_filter.value)

				__HSelect.setSelect(oElement.form.process_classes)
				__HSelect.clear(0)
				__HSelect.addArray(a,0)

				__HUtil.kill(a)
			}
		}
		else if(oElement.form.name == "oFormManagementWmi"){
			if(oElement.name == "wmi_scr_run"){
				var oSel = oElement.form.wmi_scr_language, oCod = oElement.form.wmi_scr_code
				var idx = oSel.selectedIndex, sExt = oSel.options[idx].extension, sCom = oSel.options[idx].command
				if(__H.isString(sExt,sCom)){
					var sFile = __HIO.temp + "\\mba_wmi_run." +  sExt
					if(__HFile.create(sFile,oCod.innerText)){
						oWsh.Run("%comspec% /k " + sCom + " " + sFile,__HIO.show,false)
						return;
					}
				}
				__HLog.logPopup("Unable to run script or not implemented!")
			}
			else if(oElement.name == "wmi_ope_action"){
				oDivBody.style.cursor = "wait"
				mba_common_wait("start",oElement.form.name)

				if(__H.isStringEmpty(oElement.form.server.value)) oElement.form.server.value = oWno.ComputerName
				__HMBA.wmi.u_server = oElement.form.server.value
				__HLog.logInfo("# Getting WMI namespaces for system " + __HMBA.wmi.u_server);

				if(__H.isStringEmpty(oElement.form.username.value,oElement.form.password.value)) __HMBA.wmi.u_user = __HMBA.wmi.u_pass = null // Local connections can't use password in WMI!
				else __HMBA.wmi.u_user = oElement.form.username.value, __HMBA.wmi.u_pass = oElement.form.password.value

				var aNSpaces
				__HWMI.setServiceWMI(__HMBA.wmi.u_server,null,
								__HMBA.wmi.u_user,
								__HMBA.wmi.u_pass,
								oElement.form.impersonation.value,
								oElement.form.authentication.value,
								oElement.form.privilege.value,
								oElement.form.locale.value,
								null,//oElement.form.authority.value,
								oElement.form.securityflags.value
								)

				if(aNSpaces = __HWMI.namespaces(__HMBA.wmi.u_server,null,__HMBA.wmi.u_user,__HMBA.wmi.u_pass)){
					__HSelect.setSelect(oElement.form.wmi_classes)
					__HSelect.clear(0)
					__HSelect.setSelect(oElement.form.wmi_namespace)
					__HSelect.addArray(aNSpaces,1,false,"\\\\" + __HMBA.wmi.u_server + "\\")

					management_cleartbody("oBdyManagementWmiPro")
					oElement.form.wmi_namespace.options[0].text = "=> choose a namespace.. ex: root\\cimv2"
				}
				else __HLog.popup("## Unable to make a network DCOM connection to: " + __HMBA.wmi.u_server)
				oElement.form.query.value = ""
				mba_common_wait("stop",oElement.form.name)
				oDivBody.style.cursor = "default"

				if(oElement.form.wmi_classes.options.length < 2) {
					oElement.form.wmi_namespace.selectedIndex = 1//"\\\\" + __HMBA.wmi.u_server + "\\root\\CIMV2"
					var Dummy = setTimeout("oFormManagementWmi.wmi_namespace.onchange()",0)
				}
			}
			else if(oElement.name == "wmi_namespace"){
				__HSelect.setSelect(oElement.form.wmi_namespace)
				if(!(__HMBA.wmi.u_namespace = __HSelect.getValue())) return;

				oDivBody.style.cursor = "wait"
				mba_common_wait("start",oElement.form.name);

				__HLog.logInfo("## Getting WMI classes for system " + __HMBA.wmi.u_server + " using namespace: " + __HMBA.wmi.u_namespace);
				var aOClasses
				__HWMI.setClasses(oElement.form.wmi_opt_win32.checked,oElement.form.wmi_opt_system.checked,oElement.form.wmi_opt_cim.checked)
				if(aOClasses = __HWMI.getClassesAllObject(__HMBA.wmi.u_server,__HMBA.wmi.u_namespace,__HMBA.wmi.u_user,__HMBA.wmi.u_pass)){
					__HSelect.setSelect(oElement.form.wmi_classes)
					__HSelect.clear(0)

					var len = oElement.form.wmi_classes.options.length
					for(var o in aOClasses){
						if(!aOClasses.hasOwnProperty(o)) continue
						__HSelect.addArray(aOClasses[o],len)
						len = oElement.form.wmi_classes.options.length+1
					}
					oElement.form.wmi_classes.options[0].text = "=> choose a wmi class and click 'OK'"
				}
				else __HLog.popup("## Unable to make a network DCOM connection to: " + __HMBA.wmi.u_server)
				oElement.form.query.value = ""

				mba_common_wait("stop",oElement.form.name)
				oDivBody.style.cursor = "default"

				__HUtil.kill(aOClasses)
			}
			else if(oElement.name == "wmi_classes"){
				__HSelect.setSelect(oElement.form.wmi_classes)
				if(__HMBA.wmi.u_class = __HSelect.getValue()){
					oElement.form.query.value = "Select * from " + __HMBA.wmi.u_class
					var Dummy = window.setTimeout("management_wmi_getscript('" + oElement.form.wmi_scr_language.value + "')",0);
				}
			}
			else if(oElement.name == "wmi_scr_language"){
				var Dummy = window.setTimeout("management_wmi_getscript('" + oElement.value + "')",0);
			}
		}
		else if(oElement.form.name == "oFormManagementFile"){
			var tmp, index
			if(oElement.name == "list"){
				index = parseInt(oElement.parentNode.parentNode.index)
				__HFileMan.showDirFiles(oTbyManagementFile[index],oElement.value,index)
			}
			else if(oElement.name == "browse"){
				if(tmp = __HShell.browseFolder()){
					index = parseInt(oElement.parentNode.parentNode.index)
					__HFileMan.showDirFiles(oTbyManagementFile[index],tmp,index)
				}
			}
			else if(typeof(__HFileMan[oElement.name]) == "function"){
				index = parseInt(oElement.parentNode.index)
				__HFileMan[oElement.name](oTbyManagementFile[index])
			}
		}
		return true
	}
	catch(e){
		__HLog.error(e,this);
		return false
	}
}

function management_formvalidate(oForm){
	try{

		return true
	}
	catch(e){
		__HLog.error(e,this);
		return false
	}
}

function management_navigate(sOpt,sOpt2,iIndex){
	try{
		/* Change only function name, in this case: management_navigate */

		if(!(management_navigate.caller.getName()).isSearch(/navigate/g)){
			return mba_common_navigate(sOpt,sOpt2,sOpt3)
		}

		__HMBA.mod.cur_form_exec = "EXECUTE"

		if(sOpt == "nav_management_process"){
			if(!__HMBA.mod.u_firsttime_perfom){
				ID_TIMEOUT3.push(setTimeout("oFormAction.action[3].onclick();oFormManagementProcess.process_connect.onclick()",100));
				__HMBA.mod.u_firsttime_perfom = true
			}
			mba_common_navigate("switch_this",oDivManagementProcess);
			__HMBA.mod.div2 = oDivManagementProcess1Sub

			__HMBA.mod.cur_form = oFormManagementProcess;
			__HMBA.mod.cur_form_exec = "Get Process"

			oFormAction.action[1].value = "Start"
			oFormAction.action[1].title = "Start monitoring process";
			oFormAction.action[1].onclick = function(){
				oFormManagementProcess.process_interval_stop.checked = true
				oFormManagementProcess.process_filter_action.disabled = true
				oFormManagementProcess.process_connect.disabled = true
				oFormAction.action[0].disabled = true
				__HWMI.setRefresher(__HMBA.wmi.u_class,__HMBA.wmi.u_server)
				management_process("refresh")
			}
			oFormAction.action[2].value = "Stop"
			oFormAction.action[2].title = "Stop monitoring process";
			oFormAction.action[2].className = "mba-input-br"
			oFormAction.action[2].onclick = function(){
				oFormManagementProcess.process_interval_stop.checked = false
				oFormManagementProcess.process_filter_action.disabled = false
				oFormManagementProcess.process_connect.disabled = false
				oFormAction.action[0].disabled = false
			}
			oFormAction.action[3].value = "Reset"
			oFormAction.action[3].title = "Reset " + sOpt2;
			oFormAction.action[3].onclick = new Function("mba_common_formbehavior('management_process_reset');");
		}
		else if(sOpt == "nav_management_wmi"){
			mba_common_navigate("switch_this",oDivManagementWmi);
			__HMBA.mod.cur_form_exec = "Run"
			__HMBA.mod.cur_form = oFormManagementWmi;

			oFormAction.action[1].value = "Get Methods"
			oFormAction.action[1].title = "Get WMI Methods for selected class";
			oFormAction.action[1].onclick = new Function("management_wmi('method')");
			oFormAction.action[2].value = "WBEM"
			oFormAction.action[2].title = "Start Web-Based Enterprise Management Tester";
			oFormAction.action[2].onclick = new Function("oShX.open(__HMBA.fls.wbemtest)");
			oFormAction.action[3].value = "Set Default"
			oFormAction.action[3].title = "Set Default Session " + sOpt2;
			oFormAction.action[3].onclick = new Function("mba_common_formbehavior('management_wmi_reset');");
			oFormAction.action[4].value = "Stop"
			oFormAction.action[4].title = "Stop " + sOpt2;
			oFormAction.action[4].className = "mba-input-br"
			oFormAction.action[4].onclick = new Function("mba_common_stopit(true)");
			oFormAction.action[4].disabled = true
			//oFormAction.action[4].style.backgroundColor = ""
			//oFormAction.action[4].style.filter = ""
			//oFormAction.action[4].style.marginLeft = "8px"

			if(oFormManagementWmi.wmi_namespace.options.length <= 1){
				ID_TIMEOUT1.push(setTimeout("oFormManagementWmi.wmi_ope_action.onclick()",10));
			}
		}
		else if(sOpt == "nav_management_file"){
			mba_common_navigate("switch_this",oDivManagementFile);
			__HMBA.mod.cur_form_exec = "Refresh"
			__HMBA.mod.cur_form = oFormManagementFile;

			oFormAction.action[1].value = "Reset"
			oFormAction.action[1].title = "";
			oFormAction.action[1].onclick = new Function("__HFileMan.reset()");

			if(!__HMBA.mod.u_firsttime_file){				
				setTimeout("__HFileMan.initialize()",1000)
				__HAuto.setTitle(document.title)
				__HAuto.maximize()
				oDivManagementFile_1Sup.onmousedown()
				oDivManagementFile_2Sup.onmousedown()
				__HMBA.mod.u_firsttime_file = true
			}
		}
		else return false

		oFormAction.action[0].value = "" + __HMBA.mod.cur_form_exec
		oFormAction.action[0].title = __HMBA.mod.cur_form_exec + " " + sOpt2;

		return true
	}
	catch(e){
		__HLog.error(e,this);
		return false
	}
}

function management_event(oEvent){
	try{
		if(typeof(oEvent) != "object") return;


		return;
	}
	catch(e){
		__HLog.error(e,this);
		return false
	}
}
