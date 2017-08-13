// Copyright© 2003-2008. nOsliw Solutions. All rights reserved.
//**Start Encode**

function utility_init(){
	try{
		
		return true
	}
	catch(e){
		__HLog.error(e,this);
		return false
	}
}

function utility_menu(oModule){
	try{			
		__HMBA.mnu[__HMBA.mdl_cur.menu_m] = new mnu_obj_menu(170);
		__HMBA.mnu[__HMBA.mdl_cur.menu_mm].items[__HMBA.mdl_cur.menu_i].mnu_set_sub(__HMBA.mnu[__HMBA.mdl_cur.menu_m]);		
		with(__HMBA.mnu[__HMBA.mdl_cur.menu_m]){			
			mnu_add_item(new mnu_obj_menuItem("Password Generator","","code:mba_common_navigate('nav_utility_passwd','Password Generator')"),"Ctrl-Shift-3");
			mnu_add_item(new mnu_obj_menuItem("-"));
			mnu_add_item(new mnu_obj_menuItem("Hash Cryptographic","","code:mba_common_navigate('nav_utility_hash','XmeY Hash Cryptographic')"),"Ctrl-Shift-4");
			mnu_add_item(new mnu_obj_menuItem("JScript Compression","","code:mba_common_navigate('nav_utility_compress','JScript Compression & Encryption')"),"Ctrl-Shift-5");
			mnu_add_item(new mnu_obj_menuItem("-"));
			mnu_add_item(new mnu_obj_menuItem("Molecule Formular","","code:mba_common_navigate('nav_utility_molecule','Molecule Formular')"));
			mnu_add_item(new mnu_obj_menuItem("Science & Logic","","code:mba_common_navigate('nav_utility_science','Science & Logic')"),"Ctrl-Shift-6");
		}
		
	}
	catch(e){
		__HLog.error(e,this);
		return false;
	}
}

function utility_main(){
	try{
		if(!utility_formvalidate()) return false
		
		if(__HMBA.mod.cur_form.name == "oFormUtilityPasswd"){
			return __HUtilities.password()
		}
		else if(__HMBA.mod.cur_form.name == "oFormUtilityHash"){
			if(utility_hashes_main('hash')){
				__HElem.showImg(oDivHashResultSup,oDivHashResultSup.firstChild,__HMBA.pth.pic_url + "/body/black_up_24x24.png",true);
				return true
			}
		}		
		else if(__HMBA.mod.cur_form.name == "oFormUtilityCompress"){
			return __HUtilities.compress()
		}
		else if(__HMBA.mod.cur_form.name == "oFormUtilityMolecule"){
			return __HMolecule.getFormula()
		}
		return false
	}
	catch(e){
		__HLog.error(e,this);
		return false
	}
}

function utility_formvalidate(){
	try{
		utility_navigate('nav_notimpl')
		return true
	}
	catch(e){
		__HLog.error(e,this);
		return false
	}
}

function utility_forminit(){
	try{		
		
		
		return true
	}
	catch(e){
		__HLog.error(e,this);
		return false
	}
}

function utility_formactivate(oElement){
	try{
		var oForm = oElement.form
		
		if(oForm.name == "oFormUtilityScience"){
			if(oElement.name = "quad_solve") __HScience.quadSolve()
			else if(oElement.name = "quad_clean") __HScience.quadClean()
			else if(oElement.name = "matrix_insert") __HScience.matrixInsert()
			else if(oElement.name = "matrix_solve") __HScience.matrixSolve()
		}
		return;
	}
	catch(e){
		__HLog.error(e,this);
		return false
	}
}

function utility_formbehavior(sOpt){
	try{		
		if(sOpt == "utility_passwd_reset"){
			var oForm = oFormUtilityPasswd
			oForm.reset()
			oDivUtilityPasswdResult.innerHTML = ""
		}
		else if(sOpt == "utility_hash_reset"){
			var oForm = oFormUtilityHash
			oForm.reset()
			oForm.operation1[0].onclick()
			oForm.hashes.selectedIndex = 0, oForm.hashes.onchange()
			oForm.operation2[0].onclick()
			setTimeout("utility_cleartbody('oBdyUtilityHash')",0)
		}
		else if(sOpt == "utility_compress_reset"){
			var oForm = oFormUtilityCompress
			oForm.reset()
			oDivUtilityCompressResult.innerHTML = ""
		}
		
		return true
	}
	catch(e){
		__HLog.error(e,this);
		return false
	}
}

function utility_navigate(sOpt,sOpt2,iIndex){
	try{
		if(!(utility_navigate.caller.getName()).isSearch(/navigate/g)){
			return mba_common_navigate(sOpt,sOpt2,iIndex)
		}
		
		__HMBA.mod.cur_form_exec = "Execute"
				
		if(sOpt == "nav_utility_passwd"){
			mba_common_navigate("switch_this",oDivUtilityPasswd);
			__HMBA.mod.cur_form = oFormUtilityPasswd;
			
			oFormAction.action[1].style.width = __HMBA.mod.buttom_wd2
			oFormAction.action[1].value = "Reset"
			oFormAction.action[1].title = "Reset " + sOpt2;
			oFormAction.action[1].onclick = new Function("mba_common_formbehavior('utility_passwd_reset');");
		}
		else if(sOpt == "nav_utility_hash"){
			if(!__HMBA.mod.firsttime_hash){
				__HMBA.isactive = true
				__H.temp = function(){utility_hashes_load();mba_common_formbehavior('utility_hash_reset');oDivHashXmeYSup.onmousedown();document.recalc()}
				ID_TIMEOUT3.push(window.setTimeout("__H.temp()",50));				
				__HMBA.isactive = false
				__HMBA.mod.firsttime_hash = true
			}
			mba_common_navigate("switch_this",oDivUtilityHash);
			__HMBA.mod.div2 = oDivHashResultSub			
			__HMBA.mod.cur_form = oFormUtilityHash;			
			
			oFormAction.action[1].value = "Reset"
			oFormAction.action[1].title = "Reset " + sOpt2;
			oFormAction.action[1].onclick = new Function("mba_common_formbehavior('utility_hash_reset');");
		}
		else if(sOpt == "nav_utility_compress"){
			mba_common_navigate("switch_this",oDivUtilityCompress);
			__HMBA.mod.cur_form = oFormUtilityCompress;
			
			oFormAction.action[1].value = "Reset"
			oFormAction.action[1].title = "Reset " + sOpt2;
			oFormAction.action[1].onclick = new Function("mba_common_formbehavior('utility_compress_reset');");
		}
		else if(sOpt == "nav_utility_molecule"){
			mba_common_navigate("switch_this",oDivUtilityMolecule);
			__HMBA.mod.cur_form = oFormUtilityMolecule;
			
			__HMBA.mod.cur_form_exec = "Calculate"
			
			oFormAction.action[1].value = "Add"
			oFormAction.action[1].title = "Add " + sOpt2;
			oFormAction.action[1].onclick = new Function("");
			oFormAction.action[2].value = "Substract"
			oFormAction.action[2].title = "Substract " + sOpt2;
			oFormAction.action[2].onclick = new Function("");
			oFormAction.action[3].value = "Divide"
			oFormAction.action[3].title = "Divide " + sOpt2;
			oFormAction.action[3].onclick = new Function("");
			oFormAction.action[4].value = "Multiply"
			oFormAction.action[4].title = "Multiply " + sOpt2;
			oFormAction.action[4].onclick = new Function("");
			oFormAction.action[5].value = "Reset"
			oFormAction.action[5].title = "Reset " + sOpt2;
			oFormAction.action[5].onclick = new Function("oFormUtilityMolecule.reset()");
		}
		else if(sOpt == "nav_utility_science"){
			mba_common_navigate("switch_this",oDivUtilityScience);
			__HMBA.mod.cur_form = oFormUtilityScience;
			
			__HMBA.mod.cur_form_exec = "Calculate"
			
			oFormAction.action[1].value = "Reset"
			oFormAction.action[1].title = "Reset " + sOpt2;
			oFormAction.action[1].onclick = new Function("oFormUtilityScience.reset()");
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

function utility_event(oEvent){
	try{
		if(typeof(oEvent) != "object") return;
		
		if(oEvent.ctrlKey && oEvent.shiftKey && oEvent.keyCode == 51) mba_common_navigate("nav_utility_passwd","Password Generator");
		else if(oEvent.ctrlKey && oEvent.shiftKey && oEvent.keyCode == 52) mba_common_navigate("nav_utility_hash","XmeY Hash Cryptographic");
		else if(oEvent.ctrlKey && oEvent.shiftKey && oEvent.keyCode == 53) mba_common_navigate("nav_utility_compress","JScript Compression & Encryption");
		else if(oEvent.ctrlKey && oEvent.shiftKey && oEvent.keyCode == 54) mba_common_navigate("nav_utility_science","Science & Logic");
		//else __HLog.popup("Event: " + oEvent.keyCode)
		return;
	}
	catch(e){
		__HLog.error(e,this);
		return false
	}
}