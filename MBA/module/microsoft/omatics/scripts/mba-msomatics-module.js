// Copyright© 2003-2008. nOsliw Solutions. All rights reserved.
//**Start Encode**

function msomatics_init(){
	try{
		
		
		return true
	}
	catch(e){
		__HLog.error(e,this);
		return false
	}
}

function msomatics_menu(){
	try{			
		__HMBA.mnu[__HMBA.mdl_cur.menu_m] = new mnu_obj_menu(150);		
		__HMBA.mnu[__HMBA.mdl_cur.menu_mm].items[__HMBA.mdl_cur.menu_i].mnu_set_sub(__HMBA.mnu[__HMBA.mdl_cur.menu_m]);
		with(__HMBA.mnu[__HMBA.mdl_cur.menu_m]){
			mnu_add_item(new mnu_obj_menuItem("Scripting Guys Home Page","","target:http://www.microsoft.com/technet/scriptcenter/default.mspx"),"");
			mnu_add_item(new mnu_obj_menuItem("-"));
			mnu_add_item(new mnu_obj_menuItem("Tweakomatic","","code:mba_common_navigate('nav_tweakomatic','Tweakomatic')"),"");
			mnu_add_item(new mnu_obj_menuItem("Helpomatic","","code:mba_common_navigate('nav_helpomatic','Helpomatic')"),"");
			mnu_add_item(new mnu_obj_menuItem("Scriptomatic (EzAD)","","code:mba_common_navigate('nav_adscripto','Scriptomatic (EzAD)')"),"");
		}		
	}
	catch(e){
		__HLog.error(e,this);
	}
}

function msomatics_main(){
	try{
		if(!msomatics_formvalidate()) return false
		
		if(__HMBA.mod.cur_form.name == "oFormTweakomatic"){
			return RunConfigurationScript()
		}
		else if(__HMBA.mod.cur_form.name == "oFormADScripto"){
			return RunScript()
		}
		return false
	}
	catch(e){
		__HLog.error(e,this);
		return false
	}
}

function msomatics_forminit(){
	try{
		
		
		return true
	}
	catch(e){
		__HLog.error(e,this);
		return false
	}
}

function msomatics_formbehavior(){
	try{
		
		
		return true
	}
	catch(e){
		__HLog.error(e,this);
		return false
	}
}

function msomatics_formactivate(){
	try{
		
		
		return true
	}
	catch(e){
		__HLog.error(e,this);
		return false
	}
}

function msomatics_formvalidate(){
	try{
		
		return true
	}
	catch(e){
		__HLog.error(e,this);
		return false
	}
}

function msomatics_navigate(sOpt,sOpt2,iIndex){
	try{
		if(!(msomatics_navigate.caller.getName()).isSearch(/navigate/g)){
			return mba_common_navigate(sOpt,sOpt2,iIndex)
		}
		
		__HMBA.mod.cur_form_exec = "EXECUTE"
		
		if(sOpt == "nav_tweakomatic"){
			mba_common_navigate("switch_this",oDivTweakomatic);
			
			__HMBA.mod.cur_form = oFormTweakomatic;
			__HMBA.mod.cur_form_exec = "Run"
			
			oFormAction.action[1].value = ":: Set Computer"
			oFormAction.action[1].title = "Set Computer Name";
			oFormAction.action[1].onclick = new Function("SetComputerName();");
			oFormAction.action[2].value = ":: Set Script"
			oFormAction.action[2].title = "Set Configuration Master Script";
			oFormAction.action[2].onclick = new Function("ChangeMasterFile();");
			oFormAction.action[3].value = ":: Get Script"
			oFormAction.action[3].title = "Set Retrieval Master Script";
			oFormAction.action[3].onclick = new Function("ChangeRetrievalFile();");
			
			if(!__HMBA.mdl[iIndex][sOpt + "_isfirsttime"]){
				setTimeout("Window_Onload_Tweakomatic()",0)
				__HMBA.mdl[iIndex][sOpt + "_isfirsttime"] = true
			}
		}
		else if(sOpt == "nav_adscripto"){
			mba_common_navigate("switch_this",oDivADScripto);
			
			__HMBA.mod.cur_form = oFormADScripto;
			__HMBA.mod.cur_form_exec = "Run"
			
			oFormAction.action[1].value = ":: Save"
			oFormAction.action[1].title = "";
			oFormAction.action[1].onclick = new Function("SaveScript();");
			oFormAction.action[2].value = ":: Open"
			oFormAction.action[2].title = "";
			oFormAction.action[2].onclick = new Function("OpenScript();");
			oFormAction.action[3].value = ":: Quit"
			oFormAction.action[3].title = "";
			oFormAction.action[3].onclick = new Function("QuitScript();");
			
			if(!__HMBA.mdl[iIndex][sOpt + "_isfirsttime"]){
				setTimeout("Window_Onload_ADScripto()",0)
				__HMBA.mdl[iIndex][sOpt + "_isfirsttime"] = true
			}
			else setTimeout("InitialUIState()",1000) // Must run after FormAction is created
		}
		else if(sOpt == "nav_helpomatic"){
			mba_common_navigate("switch_this",oDivHelpomatic);
			
			__HMBA.mod.cur_form = oFormHelpomatic;
			__HMBA.mod.cur_form_exec = ""
		}		
		else return false
		
		oFormAction.action[0].value = ":: " + __HMBA.mod.cur_form_exec
		oFormAction.action[0].title = __HMBA.mod.cur_form_exec + " " + sOpt2;
		
		return true
	}
	catch(e){
		__HLog.error(e,this);
		return false
	}
}

function msomatics_event(oEvent){
	try{
		
		return;
	}
	catch(e){
		__HLog.error(e,this);
		return false
	}
}
