// Copyright© 2003-2008. nOsliw Solutions. All rights reserved.
//**Start Encode**

function example_init(){
	try{
		
		return true
	}
	catch(ee){
		__HLog.error(ee,this);
		return false
	}
}

function example_menu(){
	try{
		__HMBA.mnu[__HMBA.mdl_cur.menu_m] = new mnu_obj_menu(160);		
		__HMBA.mnu[__HMBA.mdl_cur.menu_mm].items[__HMBA.mdl_cur.menu_i].mnu_set_sub(__HMBA.mnu[__HMBA.mdl_cur.menu_m]);
		with(__HMBA.mnu[__HMBA.mdl_cur.menu_m]){
			mnu_add_item(new mnu_obj_menuItem("Example SubMenu 1","i_example1","code:"));
			mnu_add_item(new mnu_obj_menuItem("-"));
			mnu_add_item(new mnu_obj_menuItem("Section Example 1","","code:mba_common_navigate('nav_example_section1','Section Example 1')"),"");
			mnu_add_item(new mnu_obj_menuItem("Section Example 2","","code:mba_common_navigate('nav_example_section2','Section Example 2')"),"");
			mnu_add_item(new mnu_obj_menuItem("Section Example 3","","code:mba_common_navigate('nav_example_section3','Section Example 3')"),"Ctrl-Alt-9");
			mnu_add_item(new mnu_obj_menuItem("-"));
			mnu_add_item(new mnu_obj_menuItem("Section Example HTML","","code:mba_common_navigate('nav_example_section4','Section Example HTML')"),"");
		}		
		
		__HMBA.mnu["m_example1"] = new mnu_obj_menu(150);
		__HMBA.mnu[__HMBA.mdl_cur.menu_m].items.i_example1.mnu_set_sub(__HMBA.mnu["m_example1"]);
		with(__HMBA.mnu["m_example1"]){
			mnu_add_item(new mnu_obj_menuItem("Item 1","","code:example_function()"),"");
			mnu_add_item(new mnu_obj_menuItem("-"));
			mnu_add_item(new mnu_obj_menuItem("Example SubMenu 2","i_example2","code:"),"");
		}
		
		
		__HMBA.mnu["m_example2"] = new mnu_obj_menu(150);
		__HMBA.mnu["m_example1"].items.i_example2.mnu_set_sub(__HMBA.mnu["m_example2"]);
		with(__HMBA.mnu["m_example2"]){
			mnu_add_item(new mnu_obj_menuItem("Go to SourceForge.Net","","target:http://www.sf.net"),"");
			mnu_add_item(new mnu_obj_menuItem("-"));
			mnu_add_item(new mnu_obj_menuItem("Open NotePad","","open:UIShell.open('notepad.exe')"),"");
		}		
	}
	catch(ee){
		__HLog.error(ee,this);
	}
}

function example_main(){
	try{
		if(!example_formvalidate()) return false
		
		if(__HMBA.mod.cur_form.name == "oFormExampleSection1"){
			return example_function2()
		}
		else if(__HMBA.mod.cur_form.name == "oFormExampleSection2"){
			return example_function2()
		}
		else if(__HMBA.mod.cur_form.name == "oFormExampleSection3"){
			return example_function2()
		}
		return false
	}
	catch(ee){
		__HLog.error(ee,this);
		return false
	}
}

function example_forminit(){
	try{
		with(oFormExampleSection1){
			tag_example1.onclick = new Function("mba_common_formactivate(this);");
			tag_example2.onclick = new Function("mba_common_formactivate(this);");
			onSubmit = new Function("return false");
		}
		
		return true
	}
	catch(ee){
		__HLog.error(ee,this);
		return false
	}
}

function example_formbehavior(sOpt){
	try{
		if(sOpt == "example_section1_reset"){
			var oForm = oFormExampleSection1;
			oForm.reset()
			
		}
		
		return true
	}
	catch(ee){
		__HLog.error(ee,this);
		return false
	}
}

function example_formactivate(oElement){
	try{
		
		if(oElement.form.name == "oFormExampleSection1") example_function()
		
		return true
	}
	catch(ee){
		__HLog.error(ee,this);
		return false
	}
}

function example_formvalidate(oForm){
	try{
		
		if(oElement.form.name == "oFormExampleSection1"){
			if(oElement.name == "tag_example1"){
				alert("you clicked button2!")
			}
		}
		
		return true
	}
	catch(ee){
		__HLog.error(ee,this);
		return false
	}
}

function example_navigate(sOpt,sOpt2,iIndex){
	try{
		/* Change only function name, in this case: example_navigate */
		
		if(!(example_navigate.caller.getName()).isSearch(/navigate/g)){
			return mba_common_navigate(sOpt,sOpt2,iIndex)
		}
		
		__HMBA.mod.cur_form_exec = "EXECUTE"
		
		/* Change below */
		
		if(sOpt == "nav_example_section1"){
			mba_common_navigate("switch_this",oDivExampleSection1);			
			__HMBA.mod.cur_form = oFormExampleSection1;
			
			oFormAction.action[1].value = "Example 1-1"
			oFormAction.action[1].title = "Title Example 1-1";
			oFormAction.action[1].onclick = new Function("example_function();");
			oFormAction.action[2].value = "Example 1-2"
			oFormAction.action[2].title = "Example " + sOpt2;
			oFormAction.action[2].onclick = new Function("example_function()");
			
			if(!__HMBA.mdl[iIndex][sOpt + "_isfirsttime"]){
				setTimeout("example_unknown_function3()",0)
				__HMBA.mdl[iIndex][sOpt + "_isfirsttime"] = true
			}
		}
		else if(sOpt == "nav_example_section2"){
			mba_common_navigate("switch_this",oDivExampleSection2);			
			__HMBA.mod.cur_form = oFormExampleSection2;
			
			__HMBA.mod.cur_form_exec = "RUN"
			
			oFormAction.action[1].value = "Example 2-1"
			oFormAction.action[1].title = "Title Example 2-1";
			oFormAction.action[1].onclick = new Function("example_unknown_function2();");
			oFormAction.action[2].value = "Example 2-2"
			oFormAction.action[2].title = "Example " + sOpt2;
			oFormAction.action[2].onclick = new Function("example_unknown_function2()");
			/*
			action: 0-4 			
			*/
		}
		else if(sOpt == "nav_example_section3"){
			mba_common_navigate("switch_this",oDivExampleSection3);
			__HMBA.mod.cur_form = oFormExampleSection3;
			
			__HMBA.mod.cur_form_exec = "OK"
			
			oFormAction.action[1].value = "Example 3-1"
			oFormAction.action[1].title = "Title Example 3-1";
			oFormAction.action[1].onclick = new Function("example_unknown_function3();");
			oFormAction.action[2].value = "Example 3-2"
			oFormAction.action[2].title = "Example " + sOpt2;
			oFormAction.action[2].onclick = new Function("example_unknown_function3()");
			/*
			action: 0-4 			
			*/
		}
		else if(sOpt == "nav_example_section4"){
			mba_common_navigate("switch_this",oDivExampleSection4);
			__HMBA.mod.cur_form = oFormExampleSection4;
			
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
	catch(ee){
		__HLog.error(ee,this);
		return false
	}
}

function example_event(oEvent){
	try{
		if(typeof(oEvent) != "object") return;
		
		if(oEvent.ctrlKey && oEvent.altKey && oEvent.keyCode == 57) mba_common_navigate('nav_example_section3','Section Example Link 3') // Ctrl-Alt-9
		
		return;
	}
	catch(ee){
		__HLog.error(ee,this);
		return false
	}
}
