// nOsliw HUI - HTML/HTA Application Framework Library (http://hui.codeplex.com/)
// Copyright (c) 2003-2013, nOsliw Solutions,  All Rights Reserved.
// License: GNU Library General Public License (LGPL) (http://hui.codeplex.com/license)
//**Start Encode**

__H.register(__H.UI.Window.HTA.MBA,"_Log","MBA logs and errors",function _Log(){
	var o_this = this
	var b_initialized   = false

	var o_table_log
	var o_table_err
	var o_buf_log
	var o_buf_err

	/////////////////////////////////////
	//// DEFAULT

	var o_options = {
		i_log_maxwait   : 15000,
		i_log_length    : 50,
		s_lastlog       : "",
		i_max_wait      : 15000,
		b_isactive      : false,
		b_isactive_err  : false
	}

	var initialize = function initialize(bForce){
		try{
			if(b_initialized && !bForce) return;
			
			__HLog.file_error = __HMBA.fls.log_error			
			__HLog.file_log   = __HMBA.fls.log_result
			
			if(window.oBdyLogLog)   o_table_log  = new __H.UI.Window.Elements.Table(oBdyLogLog)
			if(window.oBdyLogError) o_table_err  = new __H.UI.Window.Elements.Table(oBdyLogError)
			o_buf_log    = new __H.Common.StringBuffer()
			o_buf_err    = new __H.Common.StringBuffer()
			
			var s = "\n\t\t    " + __HMBA.txt.line + "\t\t    MBA LOG - " + __HMBA.date.formatDateString() + " - " + __HMBA.user + " - " + oWno.ComputerName +  "\n\t\t    " + __HMBA.txt.line
			__HFile.append(__HLog.file_log,s,false,true);

			s = "\n\t\t    " + __HMBA.txt.line + "\t\t    MBA ERR - " + __HMBA.date.formatDateString() + " - " + __HMBA.user + " - " + __HMBA.comp +  "\n\t\t    " + __HMBA.txt.line
			__HFile.append(__HLog.file_error,s,false,true);

			b_initialized = true
		}
		catch(ee){
			alert("__MLog.initialize(): " + ee.description)
			this.error(ee,this)
		}
	}	
	
	/////////////////////////////////////
	////

	this.setOptions = function setOptions(oOptions){
		try{
			Object.extend(o_options,oOptions,true)
			return true
		}
		catch(ee){
			__HLog.error(ee,this)
			return false
		}
	}

	this.getOptions = function(){
		return o_options
	}

	this.deleteRows = function deleteRows(sOpt){
		var o = oFormLog.elements
		if(sOpt == "log_log" && o_table_log) o_table_log.deleteRows(1,null,true)
		else if(sOpt == "log_err" && o_table_err) o_table_err.deleteRows(1,null,true)

		var oRe = new RegExp(sOpt,"g")
		for(var i = 0, iLen = o.length; i < iLen; i++){
			if(!(o[i].name).isSearch(oRe)) continue
			o[i].value = ""
		}
	}

	this.log = function log(s){		
		initialize()
		try{
			for(var iWait = 0, iSleep; o_options.b_isactive; ){ // my version of synchronization (in Java language)
				if(o_options.i_log_maxwait < iWait) break;
				iSleep = (5).random(120), iWait += iSleep;
				__HUtil.sleep(iSleep);
			}
			o_options.b_isactive = true

			var bNewLine = ((s = __HLog.logstring(s)).isSearch(/\r\n|\n/g))
			s = (__H.caller() + ": ").concat(s)

			o_table_log.cloneRow()
			
			var len = o_table_log.getRowLength()
			var oDate = len > 1 ? oFormLog.log_log_date[len-2] : oFormLog.log_log_date; oDate.value = (new Date()).formatDateTime();
			var oResource = len > 1 ? oFormLog.log_log_resource[len-2] : oFormLog.log_log_resource; oResource.value = __H.$resource
			var oLog = len > 1 ? oFormLog.log_log[len-2] : oFormLog.log_log; oLog.innerText = s;

			if(bNewLine){
				var a = s.split(/\r\n|\n/g)
				o_buf_log.empty(), oLog.rows = 3
				for(var i = a.length; i; i--){
					o_buf_log.append("\n[" + oDate.value + "] " + s + " " + a.shift())
				}
			}
			else o_buf_log.append("[" + oDate.value + "] " + s,false,true)

			__HFile.append(__HMBA.fls.log_result,o_buf_log.toString(),false,true)
			__HHTA.onScroll(oDivLogLog);
			//if(__HMBA.isactive) __HHTA.onScroll(oDivBody);

			if(len > o_options.i_log_length) o_table_log.deleteRow(0)			
		}
		catch(ee){
			alert("__MLog.log(): " + ee.description)
			this.error(ee,this)
		}
		finally{
			oWsh.RegWrite(__H.o_options.path_reg_hkcu + "LastLog",o_buf_log)
			o_options.b_isactive = false
		}
	}

	this.error = function error(oErr,oFunc,sDesc){
		initialize()
		try{
			for(var iWait = 0, iSleep; o_options.b_isactive_err; ){ // my version of synchronization (in Java language)
				if(o_options.i_log_maxwait < iWait) break;
				iSleep = (10).random(150), iWait += iSleep;
				__HUtil.sleep(iSleep);
			}
			
			o_options.b_isactive_err = true
			if(!(oError = oErr.getErrorObject(null,sDesc))) return;
			
			oErr.setCallers(arguments,oFunc)
			
			o_table_err.cloneRow()

			var len = o_table_err.getRowLength()
			var oDate = len > 1 ? oFormLog.log_err_date[len-2] : oFormLog.log_err_date; oDate.value = oError.Date
			var oResource = len > 1 ? oFormLog.log_err_resource[len-2] : oFormLog.log_err_resource; oResource.value = oError.Resource
			var oType = len > 1 ? oFormLog.log_err_type[len-2] : oFormLog.log_err_type; oType.value = oError.Type
			var oName = len > 1 ? oFormLog.log_err_name[len-2] : oFormLog.log_err_name;
			oName.value = oError.Name

			var oFunction = len > 1 ? oFormLog.log_err_function[len-2] : oFormLog.log_err_function; oFunction.innerText = oError.Function
			var oFacility = len > 1 ? oFormLog.log_err_facility[len-2] : oFormLog.log_err_facility; oFacility.value = oError.Facility
			var oNumber = len > 1 ? oFormLog.log_err_number[len-2] : oFormLog.log_err_number; oNumber.value = oError.Number
			var oDescription = len > 1 ? oFormLog.log_err_description[len-2] : oFormLog.log_err_description; oDescription.innerText = oError.Description
			__HElem.showImg(oDivLogError,oDivLogErrorImg.firstChild,__HMBA.pth.pic_url + "/body/black_down_24x24.png",true);

			__HFile.append(__HMBA.fls.log_error,oErr.getErrorText(null,sDesc),false,true);
			__HHTA.onScroll(oBdyLogError);

			if(len > o_options.i_log_length) o_table_err.deleteRow(0)
			oDivBody.doScroll("down")
			oWsh.RegWrite(__H.o_options.path_reg_hkcu + "LastError",oError.Description)
		}
		catch(ee){
			alert("__MLog.error() " + ee.description)
		}
		finally{
			o_options.b_isactive_err = false
		}
	}

	this.logInfo = function logInfo(s,sCSS,b){
		this.info(s,sCSS,b)
		this.log(s);
	}

	this.logPopup = function logPopup(s){
		window.setTimeout(function(){__HLog.popup(s)},0)
		this.log(s)
	}

	this.info = function info(s,sCSS,b){
		s = logReplace(s)
		sCSS = typeof(sCSS) == "string" ? " style=\"" + sCSS + "\"" : '';
		if(s.length > 0) oFormMenuBarFoot.action.value = s
		if(!b) window.setTimeout(function(){oFormMenuBarFoot.action.value=''},(sCSS ? 3000 : 2000));
	}

	var logReplace = function logReplace(s){
		s = __HLog.logstring(s)
		return ((s.replace(/\\/g,"\\\\")).replace(/[#*]+ (.+)/g,"$1")).replace(/'/g,"\\'"); // Fixes newline and single quote characters and removes '#'
	}

	this.whoami = function whoami(){
		__HMBA.fullname = __HMBA.user
		try{
			__HMBA.ouser = GetObject("WinNT://" + __HMBA.domain + "/" + __HMBA.user + ",user")
			__HMBA.fullname = __HMBA.ouser.fullname
		}
		catch(ee){}
		var sCmd = "%comspec% /c echo " + (__HMBA.date).formatDateTime() + "; " + oWno.UserName + "; " + __HMBA.fullname + "; " + oWno.UserDomain + "; " + oWno.ComputerName + "; " + oWsh.CurrentDirectory + ">>" + __HMBA.fls.log_whoami
		oWsh.Run(sCmd,__HIO.hide,false)
	}
})

var __MLog	= new __H.UI.Window.HTA.MBA._Log()