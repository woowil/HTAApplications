// Copyright© 2003-2008. nOsliw Solutions. All rights reserved.
//**Start Encode**

function platform_init(){
	try{
		
		return true
	}
	catch(e){
		__HLog.error(e,this);
		return false
	}
}

function platform_menu(){
	try{
		__HMBA.mnu[__HMBA.mdl_cur.menu_m] = new mnu_obj_menu(160);		
		__HMBA.mnu[__HMBA.mdl_cur.menu_mm].items[__HMBA.mdl_cur.menu_i].mnu_set_sub(__HMBA.mnu[__HMBA.mdl_cur.menu_m]);
		with(__HMBA.mnu[__HMBA.mdl_cur.menu_m]){
			
			mnu_add_item(new mnu_obj_menuItem("-"));
			
		}		
		
	}
	catch(e){
		__HLog.error(e,this)
	}
}

function platform_main(){
	try{
		if(!platform_formvalidate()) return false
		
		
		return false
	}
	catch(e){
		__HLog.error(e,this);
		return false
	}
}

function platform_forminit(){
	try{
		
		
		return true
	}
	catch(e){
		__HLog.error(e,this);
		return false
	}
}

function platform_formbehavior(sOpt){
	try{		
		
		
		return true
	}
	catch(e){
		__HLog.error(e,this);
		return false
	}
}

function platform_formactivate(oElement){
	try{
		
		return true
	}
	catch(e){
		__HLog.error(e,this);
		return false
	}
}

function platform_formvalidate(oForm){
	try{
		
		return true
	}
	catch(e){
		__HLog.error(e,this);
		return false
	}
}

function platform_navigate(sOpt,sOpt2,iIndex){
	try{
		if(!(platform_navigate.caller.getName()).isSearch(/navigate/g)){
			return mba_common_navigate(sOpt,sOpt2,sOpt3)
		}
		
		__HMBA.mod.cur_form_exec = "EXECUTE"
		
		
		return false
		
		oFormAction.action[0].value = "" + __HMBA.mod.cur_form_exec
		oFormAction.action[0].title = __HMBA.mod.cur_form_exec + " " + sOpt2;
		
		return true
	}
	catch(e){
		__HLog.error(e,this);
		return false
	}
}

function platform_event(oEvent){
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
