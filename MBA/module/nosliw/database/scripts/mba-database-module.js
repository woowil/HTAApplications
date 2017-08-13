// Copyright© 2003-2008. nOsliw Solutions. All rights reserved.
//**Start Encode**



function database_init(){
	try{		
		
		
		return true
	}
	catch(ee){
		__HLog.error(ee,this);
		return false
	}
}

function database_menu(){
	try{
		__HMBA.mnu[__HMBA.mdl_cur.menu_m] = new mnu_obj_menu(160);
		__HMBA.mnu[__HMBA.mdl_cur.menu_mm].items[__HMBA.mdl_cur.menu_i].mnu_set_sub(__HMBA.mnu[__HMBA.mdl_cur.menu_m]);
		with(__HMBA.mnu[__HMBA.mdl_cur.menu_m]){
			mnu_add_item(new mnu_obj_menuItem("Database Client","","code:mba_common_navigate('nav_database_client','Database Client')"),"");
			mnu_add_item(new mnu_obj_menuItem("Database Document LLD","","code:mba_common_navigate('nav_database_document_lld','Database Document LLD')"),"");
			mnu_add_item(new mnu_obj_menuItem("Database Monitor MSSQL","","code:mba_common_navigate('nav_database_monitor_ms','Database Monitor MSSQL')"),"");
		}
	}
	catch(ee){
		__HLog.error(e,this)
	}
}

function database_main(){
	try{
		if(!database_formvalidate()) return false

		if(__HMBA.mod.cur_form.name == "oFormDBClient"){
			return database_function2()
		}
		else if(__HMBA.mod.cur_form.name == "oFormDBDocumentLLD"){
			return database_function2()
		}
		else if(__HMBA.mod.cur_form.name == "oFormDBMonitorMS"){
			return database_function2()
		}
		return false
	}
	catch(ee){
		__HLog.error(ee,this);
		return false
	}
}

function database_forminit(){
	try{		
		with(oFormDBDocumentLLD){
			lld_title.onchange = new Function("mba_common_formactivate(this);");
			onSubmit = new Function("return false");
		}
	
		return true
	}
	catch(ee){
		__HLog.error(ee,this);
		return false
	}
}

function database_formbehavior(sOpt){
	try{
		if(sOpt == "database_client_reset"){
			var oForm = oFormDBClient;
			oForm.reset()

		}

		return true
	}
	catch(ee){
		__HLog.error(ee,this);
		return false
	}
}

function database_formactivate(oElement){
	try{
		if(oElement.form.name == "oFormDBClient"){
			database_function()
		}
		else if(oElement.form.name == "oFormDBDocumentLLD"){
			if(oElement.name == "lld_title"){
				__HLLD.update(oElement.value)
			}
		}
		
		return true
	}
	catch(ee){
		__HLog.error(ee,this);
		return false
	}
}

function database_formvalidate(oForm){
	try{

		return true
	}
	catch(ee){
		__HLog.error(ee,this);
		return false
	}
}

function database_navigate(sOpt,sOpt2,iIndex){
	try{
		/* Change only function name, in this case: database_navigate */

		if(!(database_navigate.caller.getName()).isSearch(/navigate/g)){
			return mba_common_navigate(sOpt,sOpt2,iIndex)
		}

		__HMBA.mod.cur_form_exec = "Run"

		/* Change below */

		if(sOpt == "nav_database_client"){
			mba_common_navigate("switch_this",oDivDBClient);
			__HMBA.mod.cur_form = oFormDBClient;

			oFormAction.action[1].value = "Client 1-1"
			oFormAction.action[1].title = "Title Client 1-1";
			oFormAction.action[1].onclick = new Function("database_function();");
			oFormAction.action[2].value = "Client 1-2"
			oFormAction.action[2].title = "Client " + sOpt2;
			oFormAction.action[2].onclick = new Function("database_function()");

			if(!__HMBA.mdl[iIndex][sOpt + "_isfirsttime"]){
				setTimeout("database_function3()",0)
				__HMBA.mdl[iIndex][sOpt + "_isfirsttime"] = true
			}
		}
		else if(sOpt == "nav_database_document_lld"){
			mba_common_navigate("switch_this",oDivDBDocumentLLD);
			__HMBA.mod.cur_form = oFormDBDocumentLLD;

			oFormAction.action[1].value = "Get LLDs"
			oFormAction.action[1].title = "Get LLD Documents";
			oFormAction.action[1].onclick = new Function("database_function();");

			if(!__HMBA.mdl[iIndex][sOpt + "_isfirsttime"]){
				setTimeout("__HLLD.init()",0)
				__HMBA.mdl[iIndex][sOpt + "_isfirsttime"] = true
			}
		}
		
		else if(sOpt == "nav_database_monitor_ms"){
			mba_common_navigate("switch_this",oDivDBMonitorMS);
			__HMBA.mod.cur_form = oFormDBMonitorMS;

			oFormAction.action[1].value = "Reset"
			oFormAction.action[1].title = "Reset all monitors";
			oFormAction.action[1].onclick = new Function("database_function();");

			if(!__HMBA.mdl[iIndex][sOpt + "_isfirsttime"]){
				//setTimeout("database_function3()",0)
				__HMBA.mdl[iIndex][sOpt + "_isfirsttime"] = true
			}
		}
		else return false

		oFormAction.action[0].value = "" + __HMBA.mod.cur_form_exec
		oFormAction.action[0].title = __HMBA.mod.cur_form_exec + " " + sOpt2;

		return true
	}
	catch(ee){
		__HLog.error(ee,this);
		return false
	}
}

function database_event(oEvent){
	try{
		if(typeof(oEvent) != "object") return;


		return ;
	}
	catch(ee){
		__HLog.error(ee,this);
		return false
	}
}
