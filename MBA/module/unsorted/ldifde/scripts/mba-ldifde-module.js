// Copyright© 2003-2008. nOsliw Solutions. All rights reserved.
//**Start Encode**

function ldifde_init(){
	try{
		
		return true
	}
	catch(e){
		__HLog.error(e,this);
		return false
	}
}

function ldifde_menu(oModule){
	try{			
		__HMBA.mnu[__HMBA.mdl_cur.menu_m] = new mnu_obj_menu(170);
		__HMBA.mnu[__HMBA.mdl_cur.menu_mm].items[__HMBA.mdl_cur.menu_i].mnu_set_sub(__HMBA.mnu[__HMBA.mdl_cur.menu_m]);
		with(__HMBA.mnu[__HMBA.mdl_cur.menu_m]){
			mnu_add_item(new mnu_obj_menuItem("LDIFDE GUI","","code:mba_common_navigate('nav_ldifde','LDIFDE')"),"");
		}		
	}
	catch(e){
		__HLog.error(e,this);
		return false;
	}
}

function ldifde_main(){
	try{
		if(!ldifde_formvalidate()) return false
		
		if(__HMBA.mod.cur_form.name == "oFormLDIFDE"){
			return LaunchCmd(__HMBA.mod.cur_form)
		}
		return false
	}
	catch(e){
		__HLog.error(e,this);
		return false
	}
}

function ldifde_formvalidate(){
	try{
		ldifde_navigate('nav_notimpl')
		return true
	}
	catch(e){
		__HLog.error(e,this);
		return false
	}
}

function ldifde_forminit(){
	try{
				
		return true
	}
	catch(e){
		__HLog.error(e,this);
		return false
	}
}

function ldifde_formactivate(oElem){
	try{
		var oForm = oElem.form
		
		return;
	}
	catch(e){
		__HLog.error(e,this);
		return false
	}
}

function ldifde_formbehavior(sOpt){
	try{		
		if(sOpt == "ldifde_reset"){
			var oForm = oFormLDIFDE;
			oForm.reset()
			
		}
		
		return true
	}
	catch(e){
		__HLog.error(e,this);
		return false
	}
}

function ldifde_navigate(sOpt,sOpt2,iIndex){
	try{
		if(!(ldifde_navigate.caller.getName()).isSearch(/navigate/g)){
			return mba_common_navigate(sOpt,sOpt2,iIndex)
		}
		
		__HMBA.mod.cur_form_exec = "Execute"
		
		if(sOpt == "nav_ldifde"){
			if(!__HMBA.mod.firsttime_ldifde){
				ID_TIMEOUT3.push(window.setTimeout("Window_Onload_LDIFDE();oDivLDIFDE2Sup.onmousedown()",10));
				__HMBA.mod.firsttime_ldifde = true
			}
			mba_common_navigate("switch_this",oDivLDIFDE);
			__HMBA.mod.cur_form = oFormLDIFDE;
			
			__HMBA.mod.cur_form_exec = "Run"
			
			oFormAction.action[1].value = "Get Syntax"
			oFormAction.action[1].title = "Generate " + sOpt2;
			oFormAction.action[1].onclick = new Function("GenerateSyntax()");			
			oFormAction.action[2].value = "Reset"
			oFormAction.action[2].title = "Reset " + sOpt2;
			oFormAction.action[2].onclick = new Function("oFormLDIFDE.reset()");
		}
		else return false
		
		oFormAction.action[0].value = "" + __HMBA.mod.cur_form_exec
		oFormAction.action[0].title = __HMBA.mod.cur_form_exec + " " + sOpt2;
		
		return true
	}
	catch(e){
		__HLog.error(e,this)
		return false;
	}
}

function ldifde_event(oEvent){
	try{
		if(typeof(oEvent) != "object") return;
		
		return;
	}
	catch(e){
		__HLog.error(e,this);
		return false
	}
}