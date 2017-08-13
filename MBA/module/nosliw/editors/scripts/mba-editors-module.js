// Copyright© 2003-2008. nOsliw Solutions. All rights reserved.
//**Start Encode**

function editors_init(){
	try{
		
		return true
	}
	catch(ee){
		__HLog.error(ee,this);
		return false
	}
}

function editors_menu(){
	try{
		__HMBA.mnu[__HMBA.mdl_cur.menu_m] = new mnu_obj_menu(160);		
		__HMBA.mnu[__HMBA.mdl_cur.menu_mm].items[__HMBA.mdl_cur.menu_i].mnu_set_sub(__HMBA.mnu[__HMBA.mdl_cur.menu_m]);
		with(__HMBA.mnu[__HMBA.mdl_cur.menu_m]){
			mnu_add_item(new mnu_obj_menuItem("Notepad","","code:mba_common_navigate('nav_editors_notepad','Notepad');"));
			mnu_add_item(new mnu_obj_menuItem("MSI Database Editor","","code:mba_common_navigate('nav_editors_msi','MSI Database Editor');"));
			
			//mnu_add_item(new mnu_obj_menuItem("-"));
			//mnu_add_item(new mnu_obj_menuItem("Manage Database","i_editors_database","code:"));
			//mnu_add_item(new mnu_obj_menuItem("Manage Table","i_editors_table","code:"));
		}		
		/*
		__HMBA.mnu["m_editors_database"] = new mnu_obj_menu(160);
		__HMBA.mnu[__HMBA.mdl_cur.menu_m].items.i_editors_database.mnu_set_sub(__HMBA.mnu["m_editors_database"]);
		with(__HMBA.mnu["m_editors_database"]){
			mnu_add_item(new mnu_obj_menuItem("New","","code:"),"");
			mnu_add_item(new mnu_obj_menuItem("Open..","","code:"),"");
			mnu_add_item(new mnu_obj_menuItem("Refresh","","code:"),"");
			mnu_add_item(new mnu_obj_menuItem("Close","","code:"),"");
			mnu_add_item(new mnu_obj_menuItem("Clear","","code:"),"");
			mnu_add_item(new mnu_obj_menuItem("-"));
			mnu_add_item(new mnu_obj_menuItem("Get Structure","","code:"),"");
			mnu_add_item(new mnu_obj_menuItem("Import/Export As Data Stream","","code:"),"");
		}		
		
		__HMBA.mnu["m_editors_table"] = new mnu_obj_menu(150);
		__HMBA.mnu[__HMBA.mdl_cur.menu_m].items.i_editors_table.mnu_set_sub(__HMBA.mnu["m_editors_table"]);
		with(__HMBA.mnu["m_editors_table"]){
			mnu_add_item(new mnu_obj_menuItem("New Table","","code:"),"");
			mnu_add_item(new mnu_obj_menuItem("New Table (custom)","","code:"),"");
			mnu_add_item(new mnu_obj_menuItem("Remove Table","","code:"),"");
			mnu_add_item(new mnu_obj_menuItem("Import Table","","code:"),"");
			mnu_add_item(new mnu_obj_menuItem("Export All Tables","","code:"),"");			
			mnu_add_item(new mnu_obj_menuItem("-"));
			mnu_add_item(new mnu_obj_menuItem("New Column","","code:"),"");			
			mnu_add_item(new mnu_obj_menuItem("-"));
			mnu_add_item(new mnu_obj_menuItem("New Row","","code:"),"");
			mnu_add_item(new mnu_obj_menuItem("Remove Row","","code:"),"");
			mnu_add_item(new mnu_obj_menuItem("-"));
			mnu_add_item(new mnu_obj_menuItem("Get Database Content","","code:"),"");
			mnu_add_item(new mnu_obj_menuItem("Edit Summary","","code:"),"");
		}
		*/		
		
	}
	catch(ee){
		__HLog.error(ee,this)
	}
}

function editors_main(){
	try{
		if(!editors_formvalidate()) return false
		
		if(__HMBA.mod.cur_form.name == "oFormEditorsNotepad"){
			return __HNotepad.close()
		}
		else if(__HMBA.mod.cur_form.name == "oFormEditorsMSI"){
			return UnLoadDB()
		}
		return false
	}
	catch(ee){
		__HLog.error(ee,this);
		return false
	}
}

function editors_forminit(){
	try{
		
		
		return true
	}
	catch(ee){
		__HLog.error(ee,this);
		return false
	}
}

function editors_formbehavior(){
	try{
		
		
		return true
	}
	catch(ee){
		__HLog.error(ee,this);
		return false
	}
}

function editors_formactivate(oElement){
	try{
		
		if(oElement.form.name == "oFormExampleSub1") editors_function()
		
		return true
	}
	catch(e){
		__HLog.error(e,this);
		return false
	}
}

function editors_formvalidate(oForm){
	try{
		
		return true
	}
	catch(e){
		__HLog.error(e,this);
		return false
	}
}

function editors_navigate(sOpt,sOpt2,iIndex){
	try{
		/* Change only function name, in this case: editors_navigate */
		
		if(!(editors_navigate.caller.getName()).isSearch(/navigate/g)){
			return mba_common_navigate(sOpt,sOpt2,sOpt3)
		}
		
		__HMBA.mod.cur_form_exec = "Close"
		
		/* Change below */
		if(sOpt == "nav_editors_notepad"){
			mba_common_navigate("switch_this",oDivEditorsNotepad);
			__HMBA.mod.cur_form = oFormEditorsNotepad;
			
			oFormAction.action[1].value = "Open"
			oFormAction.action[1].title = "";
			oFormAction.action[1].onclick = new Function("__HNotepad.open()");
			oFormAction.action[2].value = "Save"
			oFormAction.action[2].title = "";
			oFormAction.action[2].onclick = new Function("__HNotepad.save()")
			oFormAction.action[3].value = "Save As.."
			oFormAction.action[3].title = "";
			oFormAction.action[3].onclick = new Function("__HNotepad.saveas()");
			
			if(!__HMBA.mdl[iIndex][sOpt + "_isfirsttime"]){
				//setTimeout("editors_function3()",0)
				__HMBA.mdl[iIndex][sOpt + "_isfirsttime"] = true
			}
		}
		else if(sOpt == "nav_editors_msi"){
			mba_common_navigate("switch_this",oDivEditorsMSI);
			__HMBA.mod.cur_form = oFormEditorsMSI;
			
			oFormAction.action[1].value = "New"
			oFormAction.action[1].title = "";
			oFormAction.action[1].onclick = new Function("");			
			oFormAction.action[2].value = "Open"
			oFormAction.action[2].title = "";
			oFormAction.action[2].onclick = new Function("")
			oFormAction.action[3].value = "Refresh"
			oFormAction.action[3].title = "";
			oFormAction.action[3].onclick = new Function("");
			oFormAction.action[4].value = "Clear"
			oFormAction.action[4].title = "";
			oFormAction.action[4].className = "mba-input-br"
			oFormAction.action[4].onclick = new Function("");
			oFormAction.action[5].value = "Export"
			oFormAction.action[5].title = "Export Database Structure As String";
			oFormAction.action[5].onclick = new Function("");
			
			if(!__HMBA.mdl[iIndex][sOpt + "_isfirsttime"]){
				//setTimeout("editors_function3()",0)
				__HMBA.mdl[iIndex][sOpt + "_isfirsttime"] = true
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

function editors_event(oEvent){
	try{
		if(typeof(oEvent) != "object") return;
		
		if(oEvent.ctrlKey && oEvent.altKey && oEvent.keyCode == 55) mba_common_navigate('nav_example1','Section Example Link 1') // Ctrl-Alt-7
		else if(oEvent.ctrlKey && oEvent.altKey && oEvent.keyCode == 56) mba_common_navigate('nav_example2','Section Example Link 2') // Ctrl-Alt-8
		else if(oEvent.ctrlKey && oEvent.altKey && oEvent.keyCode == 57) mba_common_navigate('nav_example3','Section Example Link 3') // Ctrl-Alt-9
		
		return;
	}
	catch(e){
		__HLog.error(e,this);
		return false
	}
}
