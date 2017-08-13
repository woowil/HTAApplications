// Copyright© 2003-2008. nOsliw Solutions. All rights reserved.
//**Start Encode**


function games_init(){
	try{
		//__HMBA.mod.showIdle = false
		__HMBA.pth.pics_url = ((__HMBA.pth.cur + "\\module\\" + __HMBA.mdl_cur.module_folder + "\\data\\pic").replace(/[.]{0,1}\\/g,"/")).substr(1)
		
		__HSudoku = new __H.UI.Window.HTA.MBA._Module.Sudoku()
		ani = new __H.UI.Window.HTA.MBA._Module.GIFAnimate()
		ttt = new __H.UI.Window.HTA.MBA._Module.TicTacToe()
		
		__HSudoku.load()
		return true
	}
	catch(ee){
		__HLog.error(ee,this);
		return false
	}
}

function games_menu(){
	try{			
		__HMBA.mnu[__HMBA.mdl_cur.menu_m] = new mnu_obj_menu(150);		
		__HMBA.mnu[__HMBA.mdl_cur.menu_mm].items[__HMBA.mdl_cur.menu_i].mnu_set_sub(__HMBA.mnu[__HMBA.mdl_cur.menu_m]);
		with(__HMBA.mnu[__HMBA.mdl_cur.menu_m]){
			mnu_add_item(new mnu_obj_menuItem("Su DoKu","","code:mba_common_navigate('nav_games_sudoku','Su DoKu')"),"");
			mnu_add_item(new mnu_obj_menuItem("Tic Tac Toe","","code:mba_common_navigate('nav_games_tictactoe','Tic Tac Toe')"),"");
			mnu_add_item(new mnu_obj_menuItem("GIF Animation","","code:mba_common_navigate('nav_games_animate','GIF Animation')"),"");
		}
		
	}
	catch(ee){
		__HLog.error(ee,this);
	}
}

function games_main(){
	try{
		if(!games_formvalidate()) return false
		
		if(__HMBA.mod.cur_form.name == "oFormSudoku"){
			return __HSudoku.gameNew()
		}
		else if(__HMBA.mod.cur_form.name == "oFormAnimate"){
			return ani.init()
		}
		return false
	}
	catch(ee){
		__HLog.error(ee,this);
		return false
	}
}

function games_forminit(){
	try{
		
		
		return true
	}
	catch(ee){
		__HLog.error(ee,this);
		return false
	}
}

function games_formbehavior(){
	try{
		
		
		return true
	}
	catch(ee){
		__HLog.error(ee,this);
		return false
	}
}

function games_formactivate(oElem){
	try{
		var oForm = oElem.form
		if(oElem.form.name == "oFormSudoku"){
			if(oElem.name == "style") __HSudoku.setStyle(oElem.value)
			else if(oElem.name == "level") __HSudoku.setLevel(oElem.value)
			else if(oElem.name == "grid") __HSudoku.load(false,oElem.value)
			else if(oElem.name == "symbols") __HSudoku.setSymbols(oElem.value)
		}
		else if(oElem.form.name == "oFormAnimate"){
		
		}
		else if(oElem.form.name == "oFormTictactoe"){
		
		}
		
		return true
	}
	catch(ee){
		__HLog.error(ee,this);
		return false
	}
}

function games_formvalidate(){
	try{
		
		return true
	}
	catch(ee){
		__HLog.error(ee,this);
		return false
	}
}

function games_navigate(sOpt,sOpt2,sOpt3){
	try{
		if(!(games_navigate.caller.getName()).isSearch(/navigate/g)){
			return mba_common_navigate(sOpt,sOpt2,sOpt3)
		}
		
		__HMBA.mod.cur_form_exec = "EXECUTE"
		if(sOpt == "nav_games_sudoku"){
			mba_common_navigate("switch_this",oDivSudoku);			
			__HMBA.mod.cur_form = oFormSudoku;
			
			__HMBA.mod.cur_form_exec = "Start"
			oFormAction.action[1].value = "Solve"
			oFormAction.action[1].title = "Solve the game";
			oFormAction.action[1].onclick = new Function("__HSudoku.solve();");
			oFormAction.action[1].className = "mba-input-bb"
			oFormAction.action[2].value = "Stop"
			oFormAction.action[2].title = "Stop game time";
			oFormAction.action[2].onclick = new Function("__HSudoku.stop();");
			oFormAction.action[3].value = "Reset"
			oFormAction.action[3].title = "Reset " + sOpt2;
			oFormAction.action[3].onclick = new Function("__HSudoku.initialize();__HSudoku.clear()");
			oFormAction.action[4].value = "Clear"
			oFormAction.action[4].title = "Clear " + sOpt2;
			oFormAction.action[4].onclick = new Function("__HSudoku.clear();");
		}
		else if(sOpt == "nav_games_tictactoe"){
			mba_common_navigate("switch_this",oDivTictactoe);
			
			__HMBA.mod.cur_form = oFormTictactoe;
			__HMBA.mod.cur_form_exec = ""
			
			oFormAction.action[1].value = "Stop"
			oFormAction.action[1].title = "Stop game time";
			oFormAction.action[1].onclick = new Function("ttt.gameStop();");
			oFormAction.action[2].value = "Reset"
			oFormAction.action[2].title = "Reset " + sOpt2;
			oFormAction.action[2].onclick = new Function("ttt.setDefault();ttt.gameClear()");
			oFormAction.action[3].value = "Clear"
			oFormAction.action[3].title = "Clear " + sOpt2;
			oFormAction.action[3].onclick = new Function("ttt.gameClear();");
		}
		else if(sOpt == "nav_games_animate"){
			mba_common_navigate("switch_this",oDivAnimate);
			
			__HMBA.mod.cur_form = oFormAnimate;
			__HMBA.mod.cur_form_exec = "Load"
			
			oFormAction.action[1].value = "Slower"
			oFormAction.action[1].title = "Slower";
			oFormAction.action[1].onclick = new Function("ani.animeSlower();");
			oFormAction.action[2].value = "Stop"
			oFormAction.action[2].title = "Stop";
			oFormAction.action[2].onclick = new Function("ani.stop();");
			oFormAction.action[3].value = "Start"
			oFormAction.action[3].title = "Start " + sOpt2;
			oFormAction.action[3].onclick = new Function("ani.start(this);");
			oFormAction.action[4].value = "Faster"
			oFormAction.action[4].title = "Faster " + sOpt2;
			oFormAction.action[4].onclick = new Function("ani.animeFaster();");
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

function games_event(oEvent){
	try{
		if(typeof(oEvent) != "object") return;
		
		return;
	}
	catch(ee){
		__HLog.error(ee,this);
		return false
	}
}
