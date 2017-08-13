// Copyright© 2003-2008. nOsliw Solutions. All rights reserved.
//**Start Encode**


function mba_common_menu(){
	try{d=1
		MENU_ICON_PATH = __HMBA.pth.pic_icon_url
		// FILE
		__HMBA.mnu["m_file"] = new mnu_obj_menu(140);
		with(__HMBA.mnu["m_file"]){
			mnu_add_item(new mnu_obj_menuItem("New Frame","","code:__HShell.open(__HMBA.fls.mba_program)"),"Ctrl-N");
			mnu_add_item(new mnu_obj_menuItem("-"));
			mnu_add_item(new mnu_obj_menuItem("Print...","","code:__MDoc.printHTML(__HMBA.mod.div)"),"Ctrl-P","blue_print_16x16.png");
			mnu_add_item(new mnu_obj_menuItem("Print Frame...","","code:__MDoc.printHTML(document.body)"),"","blue_print_16x16.png");
			mnu_add_item(new mnu_obj_menuItem("-"));
			mnu_add_item(new mnu_obj_menuItem("Save As...","","code:__MDoc.saveAsHTML(__HMBA.mod.div,__HMBA.fls.div_html);"),"Ctrl-S");
			mnu_add_item(new mnu_obj_menuItem("Save Frame As...","","code:__MDoc.saveAsHTML(document.body,__HMBA.fls.div_html);"));
			mnu_add_item(new mnu_obj_menuItem("-"));
			mnu_add_item(new mnu_obj_menuItem("Refresh","","code:__HMBA.refresh()",""));
			mnu_add_item(new mnu_obj_menuItem("Reload","","code:location.reload()"),"Ctrl-R");
			mnu_add_item(new mnu_obj_menuItem("-"));
			mnu_add_item(new mnu_obj_menuItem("Exit","","code:__HHTA.closeHTA()"),"Alt-F4","logoff_20x20.gif");
			mnu_add_item(new mnu_obj_menuItem("Exit All","","code:__HHTA.closeHTA(true)"),"Alt-F5","logoff_20x20.gif");	
		}		
		d=2
		__HMBA.mnu["m_edit"] = new mnu_obj_menu(140);
		with(__HMBA.mnu["m_edit"]){
			mnu_add_item(new mnu_obj_menuItem("Undo","","code:__HWindow.setContext('undo')"),"Ctrl-Z");
			mnu_add_item(new mnu_obj_menuItem("Redo","","code:__HWindow.setContext('redo')"),"Ctrl-Y");
			mnu_add_item(new mnu_obj_menuItem("-"));
			mnu_add_item(new mnu_obj_menuItem("Cut","","code:__HWindow.setContext('cut')"),"Ctrl-X");
			mnu_add_item(new mnu_obj_menuItem("Copy","","code:__HWindow.setContext('copy')"),"Ctrl-C","copy_15x15.gif");
			mnu_add_item(new mnu_obj_menuItem("Paste","","code:__HWindow.setContext('paste')"),"Ctrl-V");
			mnu_add_item(new mnu_obj_menuItem("Delete","","code:__HWindow.setContext('delete')"),"Del","black_delete.png");
			mnu_add_item(new mnu_obj_menuItem("-"));
			mnu_add_item(new mnu_obj_menuItem("Select All","","code:__HWindow.setContext('selectall')"),"Ctrl-A");
			mnu_add_item(new mnu_obj_menuItem("-"));
			mnu_add_item(new mnu_obj_menuItem("Find...","","code:"),"Ctrl-F","black_view_16x16.png");
			mnu_add_item(new mnu_obj_menuItem("Find Next","","code:"),"F3","black_view_16x16.png");
		}		
		
		// VIEW
		__HMBA.mnu["m_view"] = new mnu_obj_menu(140);
		with(__HMBA.mnu["m_view"]){			
			mnu_add_item(new mnu_obj_menuItem("Toolbars","i_view_toolbar","code:void(0)"),"","black_edit_16x16.png");
			mnu_add_item(new mnu_obj_menuItem("Sidebars","i_view_sidebar","code:void(0)"),"","black_edit_16x16.png");
			mnu_add_item(new mnu_obj_menuItem("-"));
			mnu_add_item(new mnu_obj_menuItem("Home","","code:mba_common_navigate('nav_home','Home');"),"Alt-H");
			mnu_add_item(new mnu_obj_menuItem("Go Back","","code:mba_common_navigate('go',-1);"),"Alt-&laquo;");
			mnu_add_item(new mnu_obj_menuItem("Go Forward","","code:mba_common_navigate('go',1);"),"Alt-&raquo;");
			mnu_add_item(new mnu_obj_menuItem("Logout","","code:__HMBA.rabbit();"),"Ctrl-L");
			mnu_add_item(new mnu_obj_menuItem("-"));
			mnu_add_item(new mnu_obj_menuItem("Interface","i_view_interface","code:void(0)"));
			mnu_add_item(new mnu_obj_menuItem("Language","i_view_language","code:void(0)"));
			mnu_add_item(new mnu_obj_menuItem("-"));
			mnu_add_item(new mnu_obj_menuItem("Full Screen","","code:mba_common_navigate('go_fullscreen')"),"F11");
		}
		
		__HMBA.mnu["m_view_toolbar"] = new mnu_obj_menu(140);
		with (__HMBA.mnu["m_view_toolbar"]){
			mnu_add_item(new mnu_obj_menuItem("Menu Bar","","code:"));
			mnu_add_item(new mnu_obj_menuItem("Standard Bar","","code:"));
			mnu_add_item(new mnu_obj_menuItem("Log Bar","","code:oFormOptions.options_hideresult.onclick()"),"Alt-L","log_16x16.gif");
			mnu_add_item(new mnu_obj_menuItem("Status Bar","","code:__HElem.toggle(oDivFoot);"));
			
		}
		__HMBA.mnu["m_view"].items.i_view_toolbar.mnu_set_sub(__HMBA.mnu["m_view_toolbar"]);
		
		__HMBA.mnu["m_view_sidebar"] = new mnu_obj_menu(140);
		with (__HMBA.mnu["m_view_sidebar"]){
						
		}
		__HMBA.mnu["m_view"].items.i_view_sidebar.mnu_set_sub(__HMBA.mnu["m_view_sidebar"]);
		
		__HMBA.mnu["m_view_interface"] = new mnu_obj_menu(140);
		with (__HMBA.mnu["m_view_interface"]){
			mnu_add_item(new mnu_obj_menuItem("Set default's","","code:mba_common_setdefaults()"),"Ctrl-Alt-0");
			mnu_add_item(new mnu_obj_menuItem("-"));
			mnu_add_item(new mnu_obj_menuItem("Look & Feel","i_lnf","code:void(0)"));			
			mnu_add_item(new mnu_obj_menuItem("Colors","i_color","code:void(0)"));
			mnu_add_item(new mnu_obj_menuItem("Skins","i_skins","code:void(0)"));
		}
		__HMBA.mnu["m_view"].items.i_view_interface.mnu_set_sub(__HMBA.mnu["m_view_interface"]);
		
		__HMBA.mnu["m_lnf"] = new mnu_obj_menu(140);
		with (__HMBA.mnu["m_lnf"]){
			for(var i = 0; i < __HMBA.lnf.length; i++){
				mnu_add_item(new mnu_obj_menuItem(__HMBA.lnf[i].text_default,"","code:mba_common_setlnf(__HMBA.lnf[" + i + "]);"),"Ctrl-Alt-" + (i+1));
			}
		}
		__HMBA.mnu["m_view_interface"].items.i_lnf.mnu_set_sub(__HMBA.mnu["m_lnf"]);		

		__HMBA.mnu["m_color"] = new mnu_obj_menu(160);
		with (__HMBA.mnu["m_color"]){
			for(var i = 0; i < __HMBA.color.length; i++){
				mnu_add_item(new mnu_obj_menuItem(__HMBA.color[i].text_default,"","code:mba_common_setcolor(__HMBA.color[" + i + "]);"));
			}
		}
		__HMBA.mnu["m_view_interface"].items.i_color.mnu_set_sub(__HMBA.mnu["m_color"]);

		__HMBA.mnu["m_skins"] = new mnu_obj_menu(140);
		with (__HMBA.mnu["m_skins"]){
			for(var i = 0; i < __HMBA.skins.length; i++){
				mnu_add_item(new mnu_obj_menuItem(__HMBA.skins[i].text_default,"","code:mba_common_setskins(__HMBA.skins[" + i + "]);"));
			}
		}
		__HMBA.mnu["m_view_interface"].items.i_skins.mnu_set_sub(__HMBA.mnu["m_skins"]);
		
		__HMBA.mnu["m_view_language"] = new mnu_obj_menu(140);
		with (__HMBA.mnu["m_view_language"]){
						
		}
		__HMBA.mnu["m_view"].items.i_view_language.mnu_set_sub(__HMBA.mnu["m_view_language"]);
		
		// TOOL
		__HMBA.mnu["m_tool"] = new mnu_obj_menu(140);
		with (__HMBA.mnu["m_tool"]){
			mnu_add_item(new mnu_obj_menuItem("Mode","i_mode","code:void(0)"),"","black_config_16x16.png");
			mnu_add_item(new mnu_obj_menuItem("Programs","i_programs","code:void(0)"),"","commenu_20x20.gif");
			mnu_add_item(new mnu_obj_menuItem("-"));
			mnu_add_item(new mnu_obj_menuItem("Debug","","code:mba_common_navigate('nav_debug','Debug & Troubleshooting')"),"Alt-D","log_16x16.gif");
			mnu_add_item(new mnu_obj_menuItem("Statistics","","code:"),"")
			mnu_add_item(new mnu_obj_menuItem("-"));
			mnu_add_item(new mnu_obj_menuItem("Options","","code:mba_common_navigate('nav_options','Options')"),"Alt-O","black_config_16x16.png")
		}
		
		__HMBA.mnu["m_mode"] = new mnu_obj_menu(140);
		with (__HMBA.mnu["m_mode"]){
			for(var i = 0; i < __HMBA.mode.length; i++){
				mnu_add_item(new mnu_obj_menuItem(__HMBA.mode[i].text_default,"","code:__MCommon.setMode(__HMBA.mode[" + i + "]);"));
			}
		}
		__HMBA.mnu["m_tool"].items.i_mode.mnu_set_sub(__HMBA.mnu["m_mode"]);
		
		__HMBA.mnu["m_programs"] = new mnu_obj_menu(150);
		with(__HMBA.mnu["m_programs"]){
			mnu_add_item(new mnu_obj_menuItem("Windows Task Manager","","code:__HSys.run('taskmgr.exe')"));
			mnu_add_item(new mnu_obj_menuItem("-"));
			mnu_add_item(new mnu_obj_menuItem("Computer Management","","code:__HSys.run('compmgmt.msc /s')"));
			mnu_add_item(new mnu_obj_menuItem("System Information","","code:__HSys.run('winmsd.exe')"));
			mnu_add_item(new mnu_obj_menuItem("System Information (DOS)","","code:__HSys.getSystemInfo()"));
			mnu_add_item(new mnu_obj_menuItem("WMI Command Shell (WMIC)","","code:__HSys.run('cls & wmic.exe',true)"));
			mnu_add_item(new mnu_obj_menuItem("-"));
			mnu_add_item(new mnu_obj_menuItem("Registry Editor","","code:__HSys.run('regedit.exe')"));
			mnu_add_item(new mnu_obj_menuItem("System Configuration","","code:__HSys.run('start msconfig')"));
			mnu_add_item(new mnu_obj_menuItem("-"));
			mnu_add_item(new mnu_obj_menuItem("DirectX Diagnostic Tool","","code:__HSys.run('dxdiag.exe')"));
			mnu_add_item(new mnu_obj_menuItem("OLE/COM Object Viewer","","code:__HSys.run('oleview.exe')"));
			mnu_add_item(new mnu_obj_menuItem("ODBC Data Source Admin.","","code:__HSys.run('odbcad32.exe')"));
		}			
		__HMBA.mnu["m_tool"].items.i_programs.mnu_set_sub(__HMBA.mnu["m_programs"]);
		
		
		// MODULE
		__HMBA.mnu["m_module"] = new mnu_obj_menu(160);		
		with(__HMBA.mnu["m_module"]){
			mnu_add_item(new mnu_obj_menuItem("Manage Modules","","code:mba_common_navigate('nav_module','Manage Modules')"),"Alt-M","black_chartlink.png");
			mnu_add_item(new mnu_obj_menuItem("-"));
		}
		
		// HELP
		__HMBA.mnu["m_help"] = new mnu_obj_menu(125);
		with(__HMBA.mnu["m_help"]){
			mnu_add_item(new mnu_obj_menuItem("Help","jk","code:__MDoc.dialogHelp()"),"");
			mnu_add_item(new mnu_obj_menuItem("Help on Module","","code:mba_common_navigate('nav_docs','Help on active module')"),"F1");
			mnu_add_item(new mnu_obj_menuItem("-"));
			mnu_add_item(new mnu_obj_menuItem("Links","i_links","code:void(0)"),"","black_gallerylink_16x16.png");
			mnu_add_item(new mnu_obj_menuItem("-"));
			mnu_add_item(new mnu_obj_menuItem("About...","","code:__MDoc.dialogAbout()"),"","black_info_16x16.png");
		}		
		
		__HMBA.mnu["m_links"] = new mnu_obj_menu(160);
		with(__HMBA.mnu["m_links"]){
			mnu_add_item(new mnu_obj_menuItem("MBA on SourceForge.net","","target:http://sourceforge.net/projects/nosliw-mba/"));
			mnu_add_item(new mnu_obj_menuItem("-"));
			mnu_add_item(new mnu_obj_menuItem("HTA Applications Reference","","target:http://msdn2.microsoft.com/en-us/library/ms536473.aspx"));
			mnu_add_item(new mnu_obj_menuItem("HTA Developers Center","","target:http://www.microsoft.com/technet/scriptcenter/hubs/htas.mspx"));
			mnu_add_item(new mnu_obj_menuItem("Scripting Tools & Utilities","","target:http://www.microsoft.com/technet/scriptcenter/createit.mspx"));			
			mnu_add_item(new mnu_obj_menuItem("ASP.NET AJAX Roadmap","","target:http://asp.net/ajax/documentation/live/default.aspx"));			
			mnu_add_item(new mnu_obj_menuItem("YAHOO! UI Library: YUI Theater","","target:http://developer.yahoo.com/yui/theater/"));
		}
		__HMBA.mnu["m_help"].items.i_links.mnu_set_sub(__HMBA.mnu["m_links"]);
		
		// MENU-BAR
		__HMBA.mnu["m_menubar"] = new __HMenubar.addMenubar("static","oDivMenuBarHead","","",null); //CREATE MAIN MENU ITEMS
		with(__HMBA.mnu["m_menubar"]){
			addItem("Home",null,"home_id",true,"code:mba_common_navigate('nav_home','Home');");
			addItem("File",__HMBA.mnu["m_file"]);
			addItem("Edit",__HMBA.mnu["m_edit"]);
			addItem("View",__HMBA.mnu["m_view"]);
			addItem("Tools",__HMBA.mnu["m_tool"]);
			addItem("Modules",__HMBA.mnu["m_module"]);
			addItem("Help",__HMBA.mnu["m_help"]);
		}
		
		__HMBA.mod.menubar = __HMBA.mnu["m_menubar"]
		
		__HMBA.mnu["m_help"].items.jk.showIcon("mba-toolbar-icon-1","mba-toolbar-icon-2");
		
		// CONTEXT MENU
		__HMBA.mnu["m_context"] = new mnu_obj_menu(140);
		with(__HMBA.mnu["m_context"]){
			mnu_add_item(new mnu_obj_menuItem("Undo","","code:__HWindow.setContext('undo')"),"Ctrl-Z");
			mnu_add_item(new mnu_obj_menuItem("Redo","","code:__HWindow.setContext('redo')"),"Ctrl-Y");
			mnu_add_item(new mnu_obj_menuItem("-"));
			mnu_add_item(new mnu_obj_menuItem("Cut","","code:__HWindow.setContext('cut')"),"Ctrl-X");
			mnu_add_item(new mnu_obj_menuItem("Copy","","code:__HWindow.setContext('copy')"),"Ctrl-C");
			mnu_add_item(new mnu_obj_menuItem("Paste","","code:__HWindow.setContext('paste')"),"Ctrl-V");
			mnu_add_item(new mnu_obj_menuItem("Delete","","code:__HWindow.setContext('delete')"),"Del");
			mnu_add_item(new mnu_obj_menuItem("Append to Clipboard","","code:__HWindow.setContext('append')"));
			mnu_add_item(new mnu_obj_menuItem("-"));
			mnu_add_item(new mnu_obj_menuItem("to Upper Case","","code:__HWindow.setContext('uppercase')"));
			mnu_add_item(new mnu_obj_menuItem("to Lower Case","","code:__HWindow.setContext('lowercase')"));
			mnu_add_item(new mnu_obj_menuItem("-"));
			mnu_add_item(new mnu_obj_menuItem("Select All","","code:__HWindow.setContext('selectall')"),"Ctrl-A");
		}		
		
		// POPUP MENU
		__HMBA.mnu["m_popup"] = new mnu_obj_menu(140);
		with(__HMBA.mnu["m_popup"]){
			mnu_add_item(new mnu_obj_menuItem("Home","","code:mba_common_navigate('nav_home','Home')"),"Alt-H");
			mnu_add_item(new mnu_obj_menuItem("Logout","","code:__HMBA.rabbit();"),"Ctrl-L");
			mnu_add_item(new mnu_obj_menuItem("-"));
			mnu_add_item(new mnu_obj_menuItem("Modules","i_module","code:"),"","black_chartlink.png");
			mnu_add_item(new mnu_obj_menuItem("Close All","","code:__MModule.close()"),"Ctrl-Alt-O");
			mnu_add_item(new mnu_obj_menuItem("-"));
			mnu_add_item(new mnu_obj_menuItem("Toolbars","i_view_toolbar","code:"),"","black_edit_16x16.png");
			mnu_add_item(new mnu_obj_menuItem("Interface","i_view_interface","code:"),"");
			mnu_add_item(new mnu_obj_menuItem("-"));
			mnu_add_item(new mnu_obj_menuItem("Copy","","code:__HWindow.setContext('copy')"),"Ctrl-C","copy_15x15.gif");
			mnu_add_item(new mnu_obj_menuItem("Full Screen","","code:mba_common_navigate('go_fullscreen')"),"F11");
			mnu_add_item(new mnu_obj_menuItem("-"));			
			mnu_add_item(new mnu_obj_menuItem("Options","","code:mba_common_navigate('nav_options','Options')"),"Ctrl-O","black_config_16x16.png");
			mnu_add_item(new mnu_obj_menuItem("Help on Module","","code:mba_common_navigate('nav_docs','Help on Module')"),"F1","black_info_16x16.png");
			mnu_add_item(new mnu_obj_menuItem("-"));	
			mnu_add_item(new mnu_obj_menuItem("Refresh","","code:__HMBA.refresh()","F5"));
			mnu_add_item(new mnu_obj_menuItem("Reload","","code:location.reload()"),"Ctrl-R");
			mnu_add_item(new mnu_obj_menuItem("-"));			
			mnu_add_item(new mnu_obj_menuItem("Exit","","code:__HHTA.close()"),"Alt-F4","logoff_20x20.gif");
		}
		
		__HMBA.mnu["m_popup"].items.i_module.mnu_set_sub(__HMBA.mnu["m_module"]);
		__HMBA.mnu["m_popup"].items.i_view_toolbar.mnu_set_sub(__HMBA.mnu["m_view_toolbar"]);
		__HMBA.mnu["m_popup"].items.i_view_interface.mnu_set_sub(__HMBA.mnu["m_view_interface"]);
		
		// INITIATE
		mnu_set_popup(__HMBA.mnu["m_popup"]);
		mnu_set_activatePopupBy(0,2); //Specifies how the pop-up menu shows/hide.
		mnu_set_context(__HMBA.mnu["m_context"])
		
		oDivBody.onmousedown = mnu_set_contextMenu;
		oDivLog.onmousedown = mnu_set_contextMenu;
				
	}
	catch(e){
		__HLog.error(e,this);
		return false;
	}
	finally{
		window.setTimeout(function(){__HMBA.menu_loaded=true},1000)
	}
}

function mba_common_menuactivate(sOpt){
	try{
		if(sOpt == "init"){
			mba_common_menuactivate("edit_disable");
			mba_common_menuactivate("view_disable_go")
			return;
		}
		
		var f = __HMBA.mnu["m_file"]
		var e = __HMBA.mnu["m_edit"]
		var v = __HMBA.mnu["m_view"]
		
		if(__H.isNot(f,e,v)){
			throw new Error(8888,"The is a menu activate problem in mba_common_menuactivate()")
		}
		
		if(sOpt == "file_enable_open"){
			f.setItemsEnabled(5,6)
		}
		else if(sOpt == "file_module_close_no"){
			f.setItemsEnabled("i_modules_close")
		}
		else if(sOpt == "file_enable_close"){
			f.setItemsEnabled(7,8)
		}
		else if(sOpt == "file_delete_module_close"){
			var f_c = __HMBA.mnu["m_modules_close"]
			f_c.delItemIndex()
		}
		else if(sOpt == "edit_disable"){
			e.setItemsDisabled(0,1,3,4,5,6)
		}
		else if(sOpt == "edit_enable"){
			e.setItemsEnabled(1,3,4,5,6)
		}
		else if(sOpt == "edit_enable_paste"){
			e.setItemsEnabled(0,4)
		}
		else if(sOpt == "view_disable_go"){
			v.setItemsDisabled(2,3,"i_view_sidebar","i_view_language")
			__HMBA.mnu["m_lnf"].setItemsDisabled(2)			
		}
		else if(sOpt == "view_check_go"){return;
			if(__HMBA.mod.list.hasNext()) v.setItemsEnabled(3)
			else v.setItemsDisabled(3)
			if(__HMBA.mod.list.hasPrev()) v.setItemsEnabled(2)
			else v.setItemsDisabled(2)
		}
	}
	catch(e){
		__HLog.error(e,this);
	}
}
