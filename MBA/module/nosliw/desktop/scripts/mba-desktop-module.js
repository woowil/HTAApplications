// Copyright© 2003-2008. nOsliw Solutions. All rights reserved.
//**Start Encode**

function desktop_init(){
	try{
		
		return true
	}
	catch(e){
		__HLog.error(e,this);
		return false
	}
}

function desktop_menu(){
	try{
		__HMBA.mnu[__HMBA.mdl_cur.menu_m] = new mnu_obj_menu(160);		
		__HMBA.mnu[__HMBA.mdl_cur.menu_mm].items[__HMBA.mdl_cur.menu_i].mnu_set_sub(__HMBA.mnu[__HMBA.mdl_cur.menu_m]);
		with(__HMBA.mnu[__HMBA.mdl_cur.menu_m]){
			mnu_add_item(new mnu_obj_menuItem("Desktop SubMenu 1","i_desktop1","code:"));
			mnu_add_item(new mnu_obj_menuItem("-"));
			mnu_add_item(new mnu_obj_menuItem("Section Desktop 1","","code:mba_common_navigate('nav_desktop_section1','Section Desktop 1')"),"");
			mnu_add_item(new mnu_obj_menuItem("Section Desktop 2","","code:mba_common_navigate('nav_desktop_section2','Section Desktop 2')"),"");
			mnu_add_item(new mnu_obj_menuItem("Section Desktop 3","","code:mba_common_navigate('nav_desktop_section3','Section Desktop 3')"),"Ctrl-Alt-9");
			mnu_add_item(new mnu_obj_menuItem("-"));
			mnu_add_item(new mnu_obj_menuItem("Section Desktop HTML","","code:mba_common_navigate('nav_desktop_section4','Section Desktop HTML')"),"");
		}		
		
		__HMBA.mnu["m_desktop1"] = new mnu_obj_menu(150);
		__HMBA.mnu[__HMBA.mdl_cur.menu_m].items.i_desktop1.mnu_set_sub(__HMBA.mnu["m_desktop1"]);
		with(__HMBA.mnu["m_desktop1"]){
			mnu_add_item(new mnu_obj_menuItem("Item 1","","code:desktop_function()"),"");
			mnu_add_item(new mnu_obj_menuItem("-"));
			mnu_add_item(new mnu_obj_menuItem("Desktop SubMenu 2","i_desktop2","code:"),"");
		}
		
		
		__HMBA.mnu["m_desktop2"] = new mnu_obj_menu(150);
		__HMBA.mnu["m_desktop1"].items.i_desktop2.mnu_set_sub(__HMBA.mnu["m_desktop2"]);
		with(__HMBA.mnu["m_desktop2"]){
			mnu_add_item(new mnu_obj_menuItem("Go to SourceForge.Net","","target:http://www.sf.net"),"");
			mnu_add_item(new mnu_obj_menuItem("-"));
			mnu_add_item(new mnu_obj_menuItem("Open NotePad","","open:UIShell.open('notepad.exe')"),"");
		}		
	}
	catch(e){
		__HLog.error(e,this)
	}
}

function desktop_main(){
	try{
		if(!desktop_formvalidate()) return false
		
		if(__HMBA.mod.cur_form.name == "oFormDesktopSection1"){
			return desktop_function2()
		}
		else if(__HMBA.mod.cur_form.name == "oFormDesktopSection2"){
			return desktop_function2()
		}
		else if(__HMBA.mod.cur_form.name == "oFormDesktopSection3"){
			return desktop_function2()
		}
		return false
	}
	catch(e){
		__HLog.error(e,this);
		return false
	}
}

function desktop_forminit(){
	try{
		
		
		return true
	}
	catch(e){
		__HLog.error(e,this);
		return false
	}
}

function desktop_formbehavior(sOpt){
	try{
		if(sOpt == "desktop_section1_reset"){
			var oForm = oFormDesktopSection1;
			oForm.reset()
			
		}
		
		return true
	}
	catch(e){
		__HLog.error(e,this);
		return false
	}
}

function desktop_formactivate(oElement){
	try{
		
		if(oElement.form.name == "oFormDesktopSection1") desktop_function()
		
		return true
	}
	catch(e){
		__HLog.error(e,this);
		return false
	}
}

function desktop_formvalidate(oForm){
	try{
		
		return true
	}
	catch(e){
		__HLog.error(e,this);
		return false
	}
}

function desktop_navigate(sOpt,sOpt2,iIndex){
	try{
		/* Change only function name, in this case: desktop_navigate */
		
		if(!(desktop_navigate.caller.getName()).isSearch(/navigate/g)){
			return mba_common_navigate(sOpt,sOpt2,sOpt3)
		}
		
		__HMBA.mod.cur_form_exec = "EXECUTE"
		
		/* Change below */
		
		if(sOpt == "nav_desktop_section1"){
			mba_common_navigate("switch_this",oDivDesktopSection1);			
			__HMBA.mod.cur_form = oFormDesktopSection1;
			
			oFormAction.action[1].value = "Desktop 1-1"
			oFormAction.action[1].title = "Title Desktop 1-1";
			oFormAction.action[1].onclick = new Function("desktop_function();");
			oFormAction.action[2].value = "Desktop 1-2"
			oFormAction.action[2].title = "Desktop " + sOpt2;
			oFormAction.action[2].onclick = new Function("desktop_function()");
			
			if(!__HMBA.mdl[iIndex][sOpt + "_isfirsttime"]){
				setTimeout("desktop_function3()",0)
				__HMBA.mdl[iIndex][sOpt + "_isfirsttime"] = true
			}
		}
		else if(sOpt == "nav_desktop_section2"){
			mba_common_navigate("switch_this",oDivDesktopSection2);			
			__HMBA.mod.cur_form = oFormDesktopSection2;
			
			__HMBA.mod.cur_form_exec = "RUN"
			
			oFormAction.action[1].value = "Desktop 2-1"
			oFormAction.action[1].title = "Title Desktop 2-1";
			oFormAction.action[1].onclick = new Function("desktop_function();");
			oFormAction.action[2].value = "Desktop 2-2"
			oFormAction.action[2].title = "Desktop " + sOpt2;
			oFormAction.action[2].onclick = new Function("desktop_function()");
			/*
			action: 0-4 			
			*/
		}
		else if(sOpt == "nav_desktop_section3"){
			mba_common_navigate("switch_this",oDivDesktopSection3);
			__HMBA.mod.cur_form = oFormDesktopSection3;
			
			__HMBA.mod.cur_form_exec = "OK"
			
			oFormAction.action[1].value = "Desktop 3-1"
			oFormAction.action[1].title = "Title Desktop 3-1";
			oFormAction.action[1].onclick = new Function("desktop_function();");
			oFormAction.action[2].value = "Desktop 3-2"
			oFormAction.action[2].title = "Desktop " + sOpt2;
			oFormAction.action[2].onclick = new Function("desktop_function()");
			/*
			action: 0-4 			
			*/
		}
		else if(sOpt == "nav_desktop_section4"){
			mba_common_navigate("switch_this",oDivDesktopSection4);
			__HMBA.mod.cur_form = oFormDesktopSection4;
			
			__HMBA.mod.cur_form_exec = ""
			
			/*
			action: 0-4 
			*/
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

function desktop_event(oEvent){
	try{
		if(typeof(oEvent) != "object") return;
		
		if(oEvent.ctrlKey && oEvent.altKey && oEvent.keyCode == 57) mba_common_navigate('nav_desktop_section3','Section Desktop Link 3') // Ctrl-Alt-9
		
		return;
	}
	catch(e){
		__HLog.error(e,this);
		return false
	}
}
