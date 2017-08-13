// Copyright© 2003-2008. nOsliw Solutions. All rights reserved.
//**Start Encode**

__H.register(__H.UI.Window.HTA.MBA,"_Common","Class for window documents and dialogues",function _Common(){

	this.main = function main(){
		__HMBA.isbusy = true
		ID_TIMEOUT_MAIN = window.setTimeout(function(){__MCommon.mainThread()},10)
		//window.setTimeout('__MDoc.dialogBusy()',50)
	}

	this.mainThread = function mainThread(){
		try{
			if(__HMBA.isactive) return;
			if(__HMBA.islocked){
				__HLog.info(__HMBA.txt.locked,__HMBA.txt.locked_css);
				return;
			}

			__H.$resource = oWno.ComputerName
			mba_common_wait("start",__HMBA.mod.cur_form.name);

			if(!__HMBA.issmallscreen) __HElem.show(oDivLog);// __HElem.showImg(oDivLogLog,oDivLogLogImg.firstChild,__HMBA.pth.pic_url + "/body/black_down_24x24.png",true)

			for(var i = 0, iLen = ID_TIMEOUT3.length; i < iLen; i++) __HSys.timeStop(ID_TIMEOUT3[i]);
			var s = __HMBA.txt.line + "Title\t= " + __HMBA.title_sub + "\nUser\t= " + oWno.UserDomain + "\\" + oWno.UserName + " (" + __HMBA.fullname + ")\nHTA Version\t= " + oMBA.version + "\n";
			__HLog.log(s); // Bug... can't use log_result here.
			var bOK = window[__HMBA.mdl_cur.f_main]()
		}
		catch(e){
			__HLog.error(e,this);
			return;
		}
		finally{
			if(!__HMBA.islocked) mba_common_wait("stop",__HMBA.mod.cur_form.name);
			if(bOK){
				//__HElem.hideImg(oDivLogLog,oDivLogLogImg.firstChild,__HMBA.pth.pic_url + '/body/black_up_24x24.png',true);
				//__HElem.hideImg(oDivLogError,oDivLogErrorImg.firstChild,__HMBA.pth.pic_url + '/body/black_up_24x24.png',true);
				__HMBA.refresh()
			}
			__HMBA.isactive = false;
			__HMBA.fls.isLogExtra = __HMBA.fls.log_extra = false
			document.recalc(true);
			__HSys.timeStop(ID_TIMEOUT_MAIN)
			__HMBA.isbusy = false
		}
	}

	this.intervalCheck = function intervalCheck(){
		if(!__HMBA.isactive && !__HMBA.islocked){
			// SESSIONS
			var d = (new Date()).getDiff(__HMBA.mod.time_lasttouched);
			if(d.getMinutes() >= __HMBA.mod.time_lockedout){
				__HMBA.refresh(), __HMBA.rabbit()
			}
			else if(d.getMinutes() >= __HMBA.mod.time_lockedout-1){
				var s = "The HTA Application has been inactive in " + d.getMinutes() +
						" minutes and " + d.getSeconds() + " seconds. Session will be automatically locked after " +
						__HMBA.mod.time_lockedout + " inactive minutes"
				__HLog.popup(s,null,20000);
			}

			// PROCESSES
		}
	}

	this.intervalClock = function intervalClock(){
		var oForm = oFormMenuBarFoot
		__HMBA.isready = (!__HMBA.isactive && !__HMBA.islocked && __HMBA.isloaded)
		oForm.status.value = __HMBA.isready ? "Ready" : "Busy"
		oForm.mode.value = __HMBA.islocked ? "Locked" : __HMBA.cur_mode.text
		var d = new Date()
		if(__HMBA.mod.showIdle) oForm.idle.value = (d.getDiff(__HMBA.date)).formatHHhMMmSSs()

		oForm.time.value = d.formatHHMMSS(null,true)
		oForm.section.value = __HMBA.title_sub

		// Reloads if HTA has not been closed until the next day
		if(!__HMBA.isactive && oFormOptions.options_reload.checked){
			if(__HMBA.sdate2 != d.formatYYYYMMDD()){
				window.location.reload();
			}
		}
	}

	this.lockform = function lockform(oForm,bDisabled){
		if(typeof(oForm) == "string") oForm = __H.byIds(oForm)
		oForm.disabled = bDisabled
		
		/*
		return;
		for(var i = 0, iLen = oForm.elements.length; i < iLen; i++){
			var oElem = oForm.elements[i]
			if(oElem.className == "mba-input-br") continue
			else if((bDisabled && oElem.disabled) || (oElem.type && (oElem.type).isSearch(/radio|checkbox/ig)) || oElem.tagName == "SELECT"); // ignore tags that are already disabled
			else oElem.disabled = bDisabled;
		}*/
	}

	this.setMode = function setMode(obj){
		if(typeof(obj) != "object") return false;
		else if(__HMBA.cur_mode.index == obj.index) return true

		__HMBA.cur_mode.isActive = false
		__HMBA.cur_mode = obj

		return true
	}
})

var __MCommon = new __H.UI.Window.HTA.MBA._Common()

function mba_common_mouse(sOpt,sOpt2,oElement){
	try{
		__HMBA.mod.time_lasttouched = new Date();
		if(__HMBA.isactive && __HMBA.isloaded) return;
		else if(oElement && (oElement.id).isSearch(/oDivHomeIntro|oDivHomeSesSup|oDivRabbitSup/g));
		else if(__HMBA.islocked && __HMBA.isloaded){
			__HLog.info(__HMBA.txt.locked,__HMBA.txt.locked_css);
			return;
		}
		var Dummy = window.setTimeout(function(){mnu_hide_visable()},0);
		if(oElement.title){
			if(sOpt2 == "mouse_over") oFormMenuBarFoot.action.value = (oElement.title).toUpperCase()
			else oFormMenuBarFoot.action.value = ""
		}
		if(sOpt == "mba-input-bb"){
			if(sOpt2 == "mouse_over"){
				oElement.style.cursor = oElement.name == "help" ? "help" : "hand";
				oElement.style.color = "#000000";
				//oElement.style.backgroundColor = "#000000";
				oElement.style.filter = "progid:DXImageTransform.Microsoft.Gradient(GradientType=1,StartColorStr='#FFFFFF',EndColorStr='" + __HMBA.cur_color.col_normal + "')";
			}
			else if(sOpt2 == "mouse_out"){
				oElement.style.cursor = "default";
				//oElement.style.color = __HMBA.cur_color.col_dark
				//oElement.style.backgroundColor = "#E9967A";
				oElement.style.filter = "progid:DXImageTransform.Microsoft.Gradient(GradientType=1,StartColorStr='" + __HMBA.cur_color.col_normal + "',EndColorStr='#FFFFFF')";
				oElement.style.borderRightWidth = oElement.style.borderBottomWidth = oElement.style.borderLeftWidth = "1px";
			}
			else if(sOpt2 == "mouse_down"){
				oElement.style.backgroundColor = "#F0FFFF";
				oElement.style.color = __HMBA.cur_color.col_dark;
			}
		}
		else if(sOpt == "mba-input-bg"){
			if(sOpt2 == "mouse_over"){
				oElement.style.cursor = oElement.name == "help" ? "help" : "hand";
				oElement.style.color = "#F0FFFF";
				oElement.style.backgroundColor = "#000000";
				oElement.style.filter = "progid:DXImageTransform.Microsoft.Gradient(GradientType=1,StartColorStr='#FFFFFF',EndColorStr='#545454')";
			}
			else if(sOpt2 == "mouse_out"){
				oElement.style.cursor = "default";
				oElement.style.color = "#000000";
				oElement.style.backgroundColor = "#545454";
				oElement.style.filter = "progid:DXImageTransform.Microsoft.Gradient(GradientType=1,StartColorStr='#545454',EndColorStr='#FFFFFF')";
				oElement.style.borderRightWidth = oElement.style.borderBottomWidth = oElement.style.borderLeftWidth = "1px";
			}
			else if(sOpt2 == "mouse_down"){
				oElement.style.backgroundColor = "silver";
				oElement.style.color = __HMBA.cur_color.col_dark;
			}
		}
		else if(sOpt == "mba-input-br"){
			if(sOpt2 == "mouse_over"){
				oElement.style.cursor = "hand";
				oElement.style.color = "#FF0000";
				//oElement.style.backgroundColor = "#FF0000";
				oElement.style.filter = "progid:DXImageTransform.Microsoft.Gradient(GradientType=1,StartColorStr='#FFFFFF',EndColorStr='#E9967A')";
			}
			else if(sOpt2 == "mouse_out"){
				oElement.style.cursor = "default";
				oElement.style.color = "#FF0000";
				//oElement.style.backgroundColor = "#E9967A";
				oElement.style.filter = "progid:DXImageTransform.Microsoft.Gradient(GradientType=1,StartColorStr='#E9967A',EndColorStr='#FFFFFF')";
				oElement.style.borderRightWidth = oElement.style.borderBottomWidth = oElement.style.borderLeftWidth = "1px";
			}
			else if(sOpt2 == "mouse_down"){
				//oElement.style.backgroundColor = "#FF0000";
				oElement.style.color = __HMBA.cur_color.col_dark;
			}
		}
		else if(sOpt == "mba-div-sup"){
			if(sOpt2 == "mouse_over"){
				oElement.style.cursor = "hand";
				oElement.style.color = __HMBA.cur_color.col_normal;
				oElement.style.filter = "progid:DXImageTransform.Microsoft.Gradient(GradientType=0,StartColorStr='#FFFFFF',EndColorStr='#E4E4E4')";
			}
			else if(sOpt2 == "mouse_out"){
				oElement.style.cursor = "default";
				oElement.style.color = "#808080";
				oElement.style.filter = "progid:DXImageTransform.Microsoft.Gradient(GradientType=0,StartColorStr='#E4E4E4',EndColorStr='#FFFFFF')";
			}
			else if(sOpt2 == "mouse_down"){
				oElement.style.backgroundColor = "silver";
				oElement.style.color = __HMBA.cur_color.col_dark;

				var o = oElement.nextSibling
				if(!o || o.tagName != "DIV"){
					if(oElement.id && (oElement.id).isSearch(/Sup/g)){
						var sDiv = (oElement.id).replace(/Sup/g,"Sub");
						o = (document.getElementsByName(sDiv))[0];
					}
					else return;
				}
				if(o || o.className == "mba-div-sub"){
					__HElem.toggleImg(o,oElement.firstChild,__HMBA.pth.pic_url + "/body/black_up_16x16.png",__HMBA.pth.pic_url + "/body/black_down_16x16.png",true);
				}

				document.recalc(true)
			}
		}
		else if(sOpt == "mba-div-log"){
			if(sOpt2 == "mouse_over"){
				oElement.style.cursor = "hand";
			}
			else if(sOpt2 == "mouse_out"){
				oElement.style.cursor = "default";
			}
			else if(sOpt2 == "mouse_down"){
				var sDiv = (oElement.id).search(/error/ig) > -1 ? "oDivLogError" : "oDivLogLog";
				var aDiv = document.getElementsByName(sDiv);
				__HElem.toggleImg(aDiv[0],oElement.firstChild,__HMBA.pth.pic_url + "/body/black_down_24x24.png",__HMBA.pth.pic_url + "/body/black_up_24x24.png",true);
			}
		}
		else if(sOpt == "mba-div-arrow"){
			if(sOpt2 == "mouse_over"){
				oElement.style.cursor = "hand";
			}
			else if(sOpt2 == "mouse_out"){
				oElement.style.cursor = "default";
			}
			else if(sOpt2 == "mouse_down"){
				if(oElement.id == "oSpanActionBack") mba_common_navigate('go',-1)
				else mba_common_navigate('go',1)
			}
		}
		else if(sOpt.search(/button/ig) > -1){
			if(sOpt2 == "mouse_over"){
				oElement.style.cursor = "hand";
			}
			else if(sOpt2 == "mouse_out"){
				oElement.style.cursor = "default";
			}
		}
		else if(sOpt == "A"){
			// title is already done!
		}
		else if(sOpt == "cTR1"){
			if(sOpt2 == "mouse_over"){
				oElement.style.color = "#F0FFFF";
				oElement.bgColor = "#545454";
			}
			else if(sOpt2 == "mouse_out"){
				oElement.style.color = "#000000";
				oElement.bgColor = "#FFFFFF";
			}
		}
	}
	catch(e){
		__HLog.error(e,this);
		return false;
	}
}

function mba_common_formactivate(oElement,sOpt2){
	try{
		if(__HMBA.isactive) return;
		else if(__HMBA.islocked){
			__HLog.info(__HMBA.txt.locked,__HMBA.txt.locked_css);
			return;
		}
		var oForm = oElement.form;
		__HMBA.mod.time_lasttouched = new Date();

		if(oForm.name == "oFormBarHead"){

		}
		else if(oForm.name == "oFormMenu"){

		}
		else if(oForm.name == "oFormAction"){

		}
		else if(oForm.name == "oFormOptions"){
			if(oElement.name == "options_hideresult"){
				__HElem.toggle(oDivLog);
				oElement.checked = __HElem.isHidden(oDivLog)
			}
		}
		else if(oForm.name == "oFormDebug"){
			if(oElement.name == "dbg_script_reload1"){
				if(__HLoad.reloadScripts(oForm.dbg_script_text.value)) __HLog.popup('loaded!')
			}
			else if(oElement.name == "dbg_script_reload2"){
				if(oForm.dbg_script_list.options.length > 0){
					if(__HLoad.reloadScripts(oForm.dbg_script_list.value)) __HLog.popup('loaded!')
				}
			}
			else if((oElement.name) == "dbg_script_update"){
				var a
				if(a = __HLoad.getLoadedScripts()){
					__HSelect.setSelect(oForm.dbg_script_list)
					__HSelect.addArrayObject(a,0,"id","src");
				}
				a.empty()
			}
			else if(oElement.name == "dbg_style_reload1"){
				if(__HLoad.reloadStyle(oForm.dbg_style_text.value)) alert('loaded!')
			}
			else if(oElement.name == "dbg_style_reload2"){
				if(oForm.dbg_style_list.options.length > 0){
					if(__HLoad.reloadStyle(oForm.dbg_style_list.value)) alert('loaded!')
				}
			}
			else if((oElement.name) == "dbg_style_update"){
				var a
				if(a = __HLoad.getLoadedStyles()){
					__HSelect.setSelect(oForm.dbg_style_list)
					__HSelect.addArray(a,0);
				}
				a.empty()
			}
		}
		else if(oForm.name == "oFormModule"){
			if(oElement.name == "load"){
				eval(oForm.load_list.value)
			}
			else if(oElement.name == "loadall"){
				__MModule.loadall()
			}
			else if(oElement.name == "unload"){
				if((oForm.unload_list.value).match(/.+\(([0-9]{1,2})\)$/)){
					__MModule.unload(RegExp.$1)
				}
				else __HLog.popup("Unable to close module. Please select a module in the list.")
			}
			else if(oElement.name == "unloadall"){
				__MModule.unload()
			}
			else if(oElement.name == "create"){
				__MModule.create()
			}
		}
		else __MModule.formactivate(oElement);
	}
	catch(e){
		__HLog.error(e,this);
		return false;
	}
}

function mba_common_formbehavior(sOpt,oElement,bForce){
	try{
		if(!bForce && __HMBA.isactive && __HMBA.isloaded) return;
		var doc = document
		if(typeof(oElement) == "object" && oElement.tagName == "DIV") doc = oElement
		else if(typeof(oElement) == "object" && oElement.tagName == "FORM") var oForm = oElement, oElem = oElement.elements;

		if(sOpt == "behavior_menu"){
			for(var i = oElem.length-1; i >= 0; i--){
				var oElement = oElem[i]
				if(oElement.type == "button"){
					var c = oElement.className, n = oElement.name
					if(c == "mba-input-bb" || c == "mba-input-bg" || c == "mba-input-br"){ // Other action buttons
						oElement.title = (c.isSearch(/mba-input-bb/ig) && !n.isSearch(/help|action/ig)) ? oElement.title : oElement.value
						oElement.onmouseover = new Function("mba_common_mouse('" + c  + "','mouse_over',this);");
						oElement.onmouseout = new Function("mba_common_mouse('" + c  + "','mouse_out',this);");
						oElement.onmousedown = new Function("mba_common_mouse('" + c  + "','mouse_down',this);");
						oElement.onmouseout()
						oElement.onmousedown()
					}
				}
			}
			//__HUtil.kill(oElement)
		}
		else if(sOpt == "behavior"){
			// DYNAMIC HTML BEHAVIORS/EVENTS
			var oElement,oElements

			oElements = doc.getElementsByTagName("SELECT");
			for(var i = oElements.length-1; i >= 0; i--){
				oElements[i].onmouseover = new Function("try{this.focus()}catch(e1){}");
				if(oElements[i].onchange == null) oElements[i].onchange = new Function("mba_common_formactivate(this);")
			}

			oElements = doc.getElementsByTagName("A");
			for(var i = oElements.length-1; i >= 0; i--){
				oElements[i].onmouseover = new Function("mba_common_mouse('A','mouse_over',this);");
				oElements[i].onmouseout  = new Function("mba_common_mouse('A','mouse_out',this);");
				oElements[i].onfocus     = new Function("return false");
			}

			oElements = doc.getElementsByTagName("INPUT");
			for(var i = oElements.length-1; i >= 0; i--){
				oElement = oElements[i]
				var c = oElement.className, n = oElement.name
				if(oElement.type == "text"){ // Input text
					if(oElement.name == "server"){
						oElement.value = oWno.ComputerName
						oElement.onchange = new Function("this.value=this.value.toUpperCase();");
					}
					if(!oElement.disabled || !oElement.readonly){
						//oElement.ondblclick = new Function("this.select()")
						//oElement.onmouseover = new Function("this.style.backgroundColor='black';this.style.color='white';this.select();");
						//oElement.onmouseout = new Function("this.style.backgroundColor='#EEEEEE';this.style.color='black'");
					}
					if(c == "mba-input-tt"){
						oElement.onmouseover = new Function("mba_common_mouse('" + c  + "','mouse_over',this);");
						oElement.onmouseout = new Function("mba_common_mouse('" + c  + "','mouse_out',this);");
					}

					oElement.oncontextmenu = mnu_set_contextMenu;
				}
				else if(oElement.type == "password"){
					oElement.onpaste = new Function("oWsh.Popup('Paste events are prohibited in password fields',5,__HMBA.title_sub,48);return false;")
				}
				else if(oElement.type == "button"){
					oElement.unselectable = "on"
					if(oElement.form.name == "oFormAction");
					else if(c.isSearch(/mba-input-bb|mba-input-bg|mba-input-br/ig)){ // Other action buttons
						oElement.title = (c.isSearch(/mba-input-bb/ig) && !n.isSearch(/help|action/ig)) ? oElement.title : oElement.value
						oElement.value = ":: " + oElement.value
						oElement.onmouseover = new Function("mba_common_mouse('" + c  + "','mouse_over',this);");
						oElement.onmouseout = new Function("mba_common_mouse('" + c  + "','mouse_out',this);");
						oElement.onmousedown = new Function("mba_common_mouse('" + c  + "','mouse_down',this);");
					}
					if(oElement.onclick == null) oElement.onclick = new Function("mba_common_formactivate(this);")
				}
				else if(oElement.type == "checkbox") oElement.className = "mba-input-checkbox"
				else if(oElement.type == "radio") oElement.className = "mba-input-radio"
			}

			oElements = doc.getElementsByTagName("DIV");
			for(var c, i = oElements.length-1; i >= 0; i--){
				oElement = oElements[i]
				if(!(c = oElement.className).isSearch(/.+/));
				else if(c.isSearch(/mba-div-sup/ig)){
					oElement.unselectable = "on"
					oElement.title = oElement.innerHTML
					oElement.innerHTML = '<img src="' + __HMBA.pth.pic_url + '/body/black_down_16x16.png" hspace="4" vspace="2" align="absmiddle" border="0" alt="">' + oElement.innerHTML
					oElement.onmouseover = new Function("mba_common_mouse('" + oElement.className  + "','mouse_over',this);");
					oElement.onmouseout = new Function("mba_common_mouse('" + oElement.className  + "','mouse_out',this);");
					oElement.onmousedown = new Function("mba_common_mouse('" + oElement.className  + "','mouse_down',this);");
					oElement.onmousedown()
				}
				else if(c.isSearch(/mba-div-log/ig)){
					oElement.unselectable = "on"
					oElement.onmouseover = new Function("mba_common_mouse('" + oElement.className  + "','mouse_over',this);");
					oElement.onmouseout = new Function("mba_common_mouse('" + oElement.className  + "','mouse_out',this);");
					oElement.onmousedown = new Function("mba_common_mouse('" + oElement.className  + "','mouse_down',this);");
				}
				else if(c.isSearch(/hui-panel-frame/ig)){
					__HPanel.loadPanel(oElement)
				}
				else if(c.isSearch(/mba-div-sup/ig)){
					oElement.onmousedown()
				}
			}

			oElements = doc.getElementsByTagName("SPAN");
			for(var i = oElements.length-1; i >= 0; i--){
				oElement = oElements[i]
				var c = new String(oElement.className)
				if(c.isSearch(/mba-div-arrow/ig)){
					oElement.unselectable = "on"
					oElement.onmouseover = new Function("mba_common_mouse('" + oElement.className  + "','mouse_over',this);");
					oElement.onmouseout = new Function("mba_common_mouse('" + oElement.className  + "','mouse_out',this);");
					oElement.onmousedown = new Function("mba_common_mouse('" + oElement.className  + "','mouse_down',this);");
				}
			}

			oElements = doc.getElementsByTagName("BUTTON");
			for(var i = oElements.length-1; i >= 0; i--){
				oElement = oElements[i]
				var c = new String(oElement.className)
				if(c.isSearch(/button/ig)){
					oElement.unselectable = "on"
					oElement.onmouseover = new Function("mba_common_mouse('" + oElement.className  + "','mouse_over',this);");
					oElement.onmouseout = new Function("mba_common_mouse('" + oElement.className  + "','mouse_out',this);");
				}
			}

			oElements = doc.getElementsByTagName("LEGEND");
			for(var i = oElements.length-1; i >= 0; i--){
				if(oElements[i].parentNode.childNodes.length > 1 && (oElements[i].parentNode.childNodes(1).nodeName).isSearch(/DIV|SPAN/ig)){
					//if(oElements[i].id == "oLegRabbit") continue
					oElements[i].onmouseover = new Function("this.style.cursor='hand'");
					oElements[i].onmouseout = new Function("this.style.cursor='default'");
					oElements[i].innerHTML = '<span class="mba-span-collapse">' + __HMBA.CharExpand + '</span> ' + oElements[i].innerHTML
					// gets the DIV tag after LEGEND. gets the SPAN/DIV tag in the LEGEND,send those to to toggle()
					oElements[i].onclick = function(){
						var obj = this.parentNode.childNodes(1);
						if(__HElem.isHidden(obj)){
							__HElem.show(obj);
							this.firstChild.innerHTML = __HMBA.CharExpand
						}
						else{
							__HElem.hide(obj)
							this.firstChild.innerHTML = __HMBA.CharCollapse
						}
						document.recalc(true)
					}
				}
			}

			/*
			oElements = document.getElementsByTagName("TEXTAREA");
			for(var i = 0, iLen = oElements.length; i < iLen; i++){
				oElement = oElements[i]
				oElement.oncontextmenu = mnu_set_contextMenu;
			}
			*/
			if(!__HMBA.isfirsttimeinit){
				try{
					oDivRabbitSup.onmousedown()
					oDivHomeIntro.onmousedown()
					oFormLogin.catchmeifyoucan.focus();
				}
				catch(ee){}
			}

			//__HUtil.kill(oElements,oElement)
		}
		else if(sOpt == "options_reset"){
			oFormOptions.reset()
			oFormOptions.options_loglength.onchange()
		}
		else __MModule.formbehavior(sOpt)
	}
	catch(e){
		__HLog.popup("ERROR::mba_common_formbehavior() - " + sOpt + " " + e.description + " " +i + c)
		__HLog.error(e,this);
		return false;
	}
}



function mba_common_navigate(sOpt,sOpt2,sOpt3){
	try{
		if(__HMBA.isactive) return;

		var resetAction = function reset(){
			for(var i = oFormAction.action.length-1; i >= 0; i--){
				oFormAction.action[i].value = oFormAction.action[i].title = ""
				oFormAction.action[i].onclick = null
				oFormAction.action[i].disabled = false
				oFormAction.action[i].style.width = "auto"
			}
			if(oFormAction.help){
				oFormAction.help.onclick = null
				oFormAction.help.title = ""
			}
			oDivActionData.innerHTML = ""
			document.title = __HMBA.title = oMBA.applicationName;
		}

		if(__HMBA.islocked && __HMBA.isloaded){
			__HLog.info(__HMBA.txt.locked,__HMBA.txt.locked_css);
			return;
		}
		else if(sOpt == __HMBA.mod.div_option) return;
		else if(sOpt == "switch_all"){
			resetAction();

			__HElem.hide(__HMBA.mod.div,__HMBA.mod.div2);
			__HElem.hide(oDivOptions,oDivLog,oDivLogLog,oDivLogError,oDivDebug);
			__HElem.hideImg(oDivLogLog,oDivLogLogImg.firstChild,__HMBA.pth.pic_url + '/body/black_up_24x24.png');
			__HElem.hideImg(oDivLogError,oDivLogErrorImg.firstChild,__HMBA.pth.pic_url + '/body/black_up_24x24.png');

			return;
		}
		else if(sOpt == "switch_this"){
			if(typeof(sOpt2) != "object") return;
			__HMBA.mod.div = sOpt2,sDiv = sOpt2.id;
			__HMBA.mod.divname = (__HMBA.mod.div).id
			__HElem.show(__HMBA.mod.div);

			if(!oFormOptions.options_hideresult.checked){
				__HElem.show(oDivLog);
			}
			return;
		}
		else if(sOpt == "switch_this_alone"){
			__HMBA.mod.div = sOpt2;
			__HMBA.mod.divname = (__HMBA.mod.div).id
			__HElem.show(__HMBA.mod.div);
			__HElem.hide(oDivLog);
			return;
		}
		else if(sOpt == "go"){ // to good to be true!!
			var iIndex = (__HMBA.mod.list_index + sOpt2),oNode
			if(oNode = (__HMBA.mod.list).getNodeByIndex(null,iIndex-1)){
				mba_common_navigate(oNode.data,oNode.name,true)
				__HMBA.mod.list_index = iIndex
				//var v = __HMBA.mnu["m_view"]
				//if(sOpt < 0) (__HMBA.mod.list).hasPrev() ? v.setItemsEnabled(2) : v.setItemsDisabled(2)
				//else (__HMBA.mod.list).hasNext() ? v.setItemsEnabled(3) : v.setItemsDisabled(3)
			}
			return;
		}
		else if(sOpt == "go_fullscreen"){
			if(!__HMBA.screen.mode_3 || !__HUI.b_fullscreen){
				__HUI.screenFull()
				__HMBA.screen.mode_3 = true;
			}
			else {
				__HUI.screenRestore()
				__HMBA.screen.mode_3 = false;
			}
			//(__HMBA.mod.div).firstChild.refresh() // Refreshes section table
			return;
		}

		__HMBA.mod.time_lasttouched = new Date()
		var oForm
		__HMBA.title = __HMBA.title_sub = sOpt2

		if(sOpt.isSearch(/nav/ig)){
			mba_common_navigate('switch_all');
			var sHtml = new __H.Common.StringBuffer('<table width="100%" align="center" border="0" cellspacing="0" cellpadding="0"><tr><td class="mba-title-2Text"><div id="oDivTitle2Text" class="mba-div-filter">\u25A1 ' + sOpt2 + '</div></td>')
			sHtml.append('<td align="right"><input name="info" type="hidden">')
			if(!sOpt.isSearch(/nav_home|nav_docs/ig)){
				for(var i = 0; i < __HMBA.mod.action_len; i++){
					var sInput = i == 0 ? "mba-input-bb" : "mba-input-bg"
					var sStyle = i == 0 ? "background-color:" + __HMBA.cur_color.col_normal + ";color:" + __HMBA.cur_color.col_dark + ";border-color:" + __HMBA.cur_color.col_dark + ";border-style:solid;border-width: 2px 1px 1px 1px" : "color: #000000;"
					sHtml.append('<input type="button" title="" value="" name="action" class="' + sInput + '" style="width:auto;' + sStyle + '"> ')
					sHtml.append((i == 0 ? " &nbsp; " : ""))
				}
			}
			sHtml.append(' &nbsp;&nbsp; </td><td width="1%"> <input type="button" value=" ? " name="help" class="mba-input-bb" style="width:auto;border-color:' + __HMBA.cur_color.col_dark + ';" onmouseover="this.style.cursor=\'help\'"></td></tr></table>')
			oDivActionData.innerHTML = sHtml.toString()
			__HLog.info("# Navigating to - " + __HMBA.title_sub);
		}

		if(sOpt == "nav_notimpl"){
			__HLog.info(__HMBA.txt.notimpl);
			return;
		}
		else if(sOpt == "nav_home"){
			mba_common_navigate("switch_this_alone",oDivHome);
			__HMBA.mod.cur_form = oFormHome;
		}
		else if(sOpt == "nav_docs"){
			mba_common_navigate("switch_this_alone",oDivDocs);
			__HMBA.mod.cur_form = oFormDocs;
			if(__HMBA.mod.firsttime_docs){
				__HMBA.mod.firsttime_docs = !__MDoc.getDocs();
			}
		}
		else if(sOpt == "nav_debug"){
			mba_common_navigate("switch_this_alone",oDivDebug);
			__HMBA.mod.cur_form = oFormDebug;

			oFormAction.action[1].value = "Reset"
			oFormAction.action[1].onclick = new Function("oFormDebug.reset()");

			if(!__HMBA.mod.firsttime_debug){
				//__HMBA.mod.firsttime_options = mba_common_options("config");
				__HMBA.mod.firsttime_debug = true
			}
			__HElem.show(oDivLog)
		}
		else if(sOpt == "nav_options"){
			mba_common_navigate("switch_this_alone",oDivOptions);
			__HMBA.mod.cur_form = oFormOptions;

			oFormAction.action[1].value = oFormAction.action[1].title = "Reset";
			oFormAction.action[1].style.width = __HMBA.mod.buttom_wd2
			oFormAction.action[1].onclick = new Function("mba_common_formbehavior('options_reset');");

			if(!__HMBA.mod.firsttime_options){
				//__HMBA.mod.firsttime_options = mba_common_options("config");
				window.setTimeout(function(){__HPanel.loadPanel(oDivOptionsPanel)},0)
				__HMBA.mod.firsttime_options = true
			}
		}
		else if(sOpt == "nav_module"){
			mba_common_navigate("switch_this_alone",oDivModule);
			__HMBA.mod.cur_form = oFormModule;

			oFormAction.action[1].value = "Initialize"
			oFormAction.action[1].className = "mba-input-bb"
			oFormAction.action[1].onclick = new Function("__MModule.initialize()");
			oFormAction.action[1].title = "Initializes/updates the module load list"
			oFormAction.action[2].value = "Reset"
			oFormAction.action[2].onclick = new Function("oFormModule.reset();__MModule.busy(false)");

			if(!__HMBA.mod.firsttime_module){
				window.setTimeout(function(){__HElem.transition(oDivModuleImg)},250)
				if(oFormModule.load_list.options.length < 2){
					window.setTimeout(function(){oFormAction.action[0].onclick()},50)
				}
				__HMBA.mod.firsttime_module = true
			}
			__HElem.showImg(oDivModule1Sub,oDivModule1Sup.firstChild,__HMBA.pth.pic_url + "/mba_arrow2_up.gif",true);
		}
		else if(__MModule.navigate(sOpt,sOpt2,sOpt3));
		else return;

		if(!sOpt.isSearch(/nav_home|nav_docs/ig)){
			oFormAction.info.value = __HMBA.mod.cur_form.name
			oFormAction.action[0].onclick = new Function("__MCommon.main()");
			oFormAction.help.title = "Help(F1) on " + sOpt2
			oFormAction.help.onclick = new Function("__MDoc.dialogHelp()")

			for(var i = oFormAction.action.length-1; i >= 0; i--){
				var o = oFormAction.action[i], c = o.className
				if((o.value).length > 0){
					if(!o.title) o.title = o.value + " " + sOpt2;
					o.value = ":: " + o.value
					o.style.width = __HMBA.mod.buttom_wd2
				}
				else o.removeNode(true)
			}
		}

		if(sOpt.isSearch(/nav/ig)) mba_common_formbehavior("behavior_menu",oFormAction,true)
		if(!__HMBA.mod.list) __HMBA.mod.list = new __H.Util.List.LinkedDouble(sOpt,sOpt2), __HMBA.mod.list_index = 1;
		else (__HMBA.mod.list).addTail(null,sOpt,sOpt2)

		if(!sOpt3) __HMBA.mod.list_index = (__HMBA.mod.list).getSize(); // Only when not using arrow
		__HMBA.mod.div_option = sOpt

		window.setTimeout(function(){__HElem.transition(oDivTitle2Text)},__HMBA.mod.fadedelay1)
		return;
	}
	catch(e){
		__HLog.error(e,this);
		return;
	}
	finally{
		document.recalc(true);
		window.setTimeout(function(){mnu_hide_context()},0)
		window.setTimeout(function(){mba_common_menuactivate('view_check_go')},0)
	}
}

function mba_common_event(){
	try{
		var oEvent = window.event;
		if(__HMBA.isactive) return;
		__HMBA.mod.time_lasttouched = new Date();

		if(arguments[0] == "init" || (oEvent.altKey && oEvent.keyCode == 87)){ // oops I did it again!
			if(!__HMBA.isloaded) return;
			__HMBA.islocked = __HMBA.islocked ? false : true;
			
			if(!__HMBA.isfirsttimeinit){				
				mba_common_navigate("nav_home","Home")
				__HMBA.refresh()				
				mba_common_menuactivate('init')
				// Clean all none wanted code in MBA. Must follow rules using DIV				
				for(var i = oDivBody.childNodes.length-1; i >= 0; i--){
					var o = oDivBody.childNodes(i)
					if(o.tagName != "DIV"){
						__HElem.hide(o)
					}
				}
				__HHTA.position(950,(window.screen.height*0.9))
				__HMBA.isfirsttimeinit = true
			}
			__HElem.hide(oDivLogin,oDivHomeSplash)
			__HElem.show(oTblTop)		
			window.setTimeout(function(){mba_common_navigate('go_fullscreen')},50)
			window.setTimeout(function(){__HElem.transition(oDivHomeSplash)},500)
		}
		else if(oEvent.ctrlKey && oEvent.keyCode == 82) window.location.reload(); //
		else if(oEvent.altKey && oEvent.keyCode == 115) __HHTA.closeHTA(); // Alt-F4
		else if(oEvent.altKey && oEvent.keyCode == 116) __HHTA.closeHTA(true); // Alt-F5
		else if(oEvent.keyCode == 27) mba_common_stop(true,oEvent); // ESC
		else if(oEvent.keyCode == 112){ // F1 / Help
			try{
				oFormAction.help.onclick();
			}
			catch(ee){}
		}
		else if(oEvent.altKey && oEvent.keyCode == 82) __HElem.toggle(oDivLog); // Alt R
		else if(oEvent.keyCode == 13){
			var oElem = oEvent.srcElement
			if(!oElem || oElem.tagName != "INPUT") return;
			var sName = "9fo9kjbnaAtr9fi".encode();

			if(oElem.name == sName){
				if(oElem.form[sName].value != oFormMenu[__HMBA.mod.checked].value){
					__HLog.popup(__HMBA.txt.nocando);
					return false;
				}
				oElem.form[sName].disabled = true
				mba_common_event("init");
				__HLog.info(__HMBA.txt.unlocked);
			}
		}
		else if(oEvent.ctrlKey && oEvent.keyCode == 78) __HShell.open(__HMBA.fls.mba_program);// Ctrl-N
		else if(__HMBA.islocked && oEvent.keyCode != 17 && oEvent.keyCode != 18) __HLog.info(__HMBA.txt.locked);
		else if(oEvent.ctrlKey && oEvent.keyCode == 80) __MDoc.printHTML(__HMBA.mod.div); // Ctrl-P
		else if(oEvent.ctrlKey && oEvent.keyCode == 83) __MDoc.saveAsText(__HMBA.mod.div,__HMBA.fls.div_html);
		else if(oEvent.altKey && oEvent.keyCode == 79) mba_common_navigate("nav_options","Options");

		else if(oEvent.altKey && oEvent.keyCode == 68) mba_common_navigate("nav_debug","Debug & Troubleshooting")
		else if(oEvent.altKey && oEvent.keyCode == 72) mba_common_navigate("nav_home","Home");
		else if(oEvent.altKey && oEvent.keyCode == 37) mba_common_navigate('go',-1);
		else if(oEvent.altKey && oEvent.keyCode == 39) mba_common_navigate('go',1);
		else if(oEvent.ctrlKey && oEvent.keyCode == 76) __HMBA.rabbit() // Ctrl-L

		else if(oEvent.altKey && oEvent.keyCode == 76) oFormOptions.options_hideresult.onclick() // Alt-L
		else if(oEvent.altKey && oEvent.keyCode == 77) mba_common_navigate('nav_module','Manage Modules') // Alt-M
		else if(oEvent.keyCode == 122) mba_common_navigate("go_fullscreen");

		else if(oEvent.ctrlKey && oEvent.altKey && oEvent.keyCode == 49) mba_common_setlnf(__HMBA.lnf[0]); // Ctrl-Alt-1
		else if(oEvent.ctrlKey && oEvent.altKey && oEvent.keyCode == 50) mba_common_setlnf(__HMBA.lnf[1]); // Ctrl-Alt-2
		else if(oEvent.ctrlKey && oEvent.altKey && oEvent.keyCode == 51) mba_common_setlnf(__HMBA.lnf[2]); // Ctrl-Alt-3
		else if(oEvent.ctrlKey && oEvent.keyCode == 65) __HMBA.mod.div.document.execCommand("SelectAll"); // Ctrl-A
		else return __MModule.event(oEvent)
		// http://www.htmlhelp.com/reference/charset/iso032-063.html
		//else if(!__HMBA.islocked) __HLog.popup("Event: " + oEvent.keyCode)
	}
	catch(e){
		__HLog.error(e,this);
		return;
	}
}


function mba_common_setdefaults(obj){
	try{
		mba_common_setlnf(__HMBA.lnf[0])
		mba_common_setcolor(__HMBA.color[0])
		mba_common_setskins(__HMBA.skins[0])
	}
	catch(e){
		__HLog.error(e,this);
		return false;
	}
}

function mba_common_setlnf(obj){
	try{
		if(typeof(obj) != "object") return false;
		else if(__HMBA.cur_lnf.index == obj.index) return true
		else if((obj.text).isSearch(/tree/ig)){
			__HLog.info(__HMBA.txt.notimpl);
			return true;
		}

		window.setTimeout(function(){mnu_hide_context()},0)
		__HElem.hide(oDivMenuBarHead,oDivMenuBarFoot,oDivListBarHead,oDivListBarFoot,oDivList1,oDivList2)

		if((obj.text).isSearch(/menu/ig)){
			__HElem.show(oDivMenuBarHead,oDivMenuBarFoot)

			if(oDivActionLoc2.innerText.length > 2){
				oDivActionLoc1.innerHTML = oDivActionLoc2.innerHTML
				oDivActionLoc2.innerText = ""
			}
		}
		else if((obj.text).isSearch(/list/ig)){
			if(!__HMBA.mod.isListCreated){
				// Menu
				oDivListBarHead.innerHTML = (oDivListBarHead.innerHTML).concat(__HMBA.txt.mba_info)
				oDivListBarFoot.innerHTML = __HMBA.txt.mba_info + oDivListBarFoot.innerHTML
				__HElem.show(oDivListBarHead,oDivListBarFoot)
				// List
				var oSel = __H.byClone("select")
				var o = oDivMenuBarHead.childNodes
				for(var i = 0, iLen = o.length; i < iLen; i++){
					if((o[i].innerHTML).length < 2) continue
					var oOpt = __H.byClone("option")
					oOpt.value = o[i].innerHTML
					oOpt.innerHTML = " &#9633; " + o[i].innerHTML + " &raquo;"
					oSel.appendChild(oOpt)
				}
				oSel.onchange = new Function("mba_common_formactivate(this)")
				oSel.style.marginLeft = "5px"
				oDivList1.appendChild(oSel)

				if(oDivActionLoc1.innerText.length > 2){
					oDivActionLoc2.innerHTML = oDivActionLoc1.innerHTML
					oDivActionLoc1.innerText = ""
				}

				__HMBA.mod.isListCreated = true
			}
			__HElem.show(oDivList1,oDivList2)
		}

		__HMBA.cur_lnf.isActive = false
		__HMBA.cur_lnf = obj

		return true
	}
	catch(e){
		__HLog.error(e,this);
		return false;
	}
	finally{
		oTblTop.refresh()
		document.recalc(true)
	}
}

function mba_common_setcolor(obj){
	try{
		if(typeof(obj) != "object") return false;
		else if(__HMBA.cur_color.index == obj.index) return true
		var oElements, oElement

		__MStyle.setMenuColor((__HMBA.mod.menubar).menuBarObj,obj.col_dark,obj.col_normal)

		//var style = __HStyle.getStyle(".mba-input-bb")
		__HStyle.setStyle(".mba-input-bb","backgroundColor",obj.col_normal)
		__HStyle.setStyle(".mba-input-bb","borderColor",obj.col_dark)
		__HStyle.setStyle(".mba-input-bb","filter","progid:DXImageTransform.Microsoft.Gradient(GradientType=1,StartColorStr='" + obj.col_normal + "',EndColorStr='#FFFFFF')")
		__HStyle.setStyle(".mba-about-1","backgroundColor",obj.col_normal)
		__HStyle.setStyle(".mba-about-1","filter","progid:DXImageTransform.Microsoft.Gradient(GradientType=0,StartColorStr='" + obj.col_normal + "',EndColorStr='#FFFFFF')")
		//__HStyle.setStyle(".mba-div-sup","color",obj.col_dark)
		__HStyle.setStyle(".mba-span-note","color",obj.col_dark)
		__HStyle.setStyle("LEGEND","color",obj.col_dark)
		__HStyle.setStyle("A","color",obj.col_dark)
		__HStyle.setStyle("A:link","color",obj.col_dark)
		__HStyle.setStyle("A:visited","color",obj.col_normal)
		__HStyle.setStyle("A:active","color",obj.col_normal)

		__HMBA.txt.notimpl_css =  "color:" + obj.col_dark
		__HMBA.txt.info_css = "color:" + obj.col_dark + ";font-size:12";
		__HMBA.txt.credentials_css = __HMBA.txt.info_css
		__HMBA.txt.deactivated_css = __HMBA.txt.info_css

		//document.linkColor = obj.col_dark
		//document.alinkColor = document.vlinkColor = obj.col_normal
		__HMBA.cur_color.isActive = false

		obj.isActive = true
		__HMBA.cur_color = obj

		mba_common_formbehavior("behavior_menu",oFormAction)
	}
	catch(e){
		__HLog.error(e,this);
		return false;
	}
}

function mba_common_setskins(obj){
	try{
		if(typeof(obj) != "object") return false;
		else if(__HMBA.cur_skins.index == obj.index) return true

		if(__HLoad.loadStyle(obj.css_href)){
			oDivMenuBarFoot.firstChild.bgColor = obj.col_statusbar
			__HMBA.cur_skins.isActive = false
			obj.isActive = true
			__HMBA.cur_skins = obj
			document.recalc(true)
			// Remove other styles
			for(var i = 0; i < __HMBA.skins[0].length; i++){
				if(i == obj.index) continue
				__HLoad.removeStyle(__HMBA.skins[0].css_href)
			}
			return true
		}

		return false;
	}
	catch(e){
		__HLog.error(e,this);
		return false;
	}
}

function mba_common_stopit(bStop,ev){
	try{ // http://www.zytrax.com/tech/dom/guide.html#events__H
		if(bStop && !confirm("This will invoke a safe stop on any running activity. Continue?")) return false;
		try{
			var ev = ev ? ev : window.event
			if(ev.button & 2){ // Note: We use & not == in the tests for the right button in case multiple buttons were pressed.
				if(ev.stopPropagation){
					ev.stopPropagation(); // prevent right click context menu
				}
			}
			ev.cancelBubble = true;
			return true;
		}
		catch(ee){}

		for(var i = 0, iLen = oFormAction.action.length; i < iLen; i++){
			if(oFormAction.action[i].className == "mba-input-br"){
				oFormAction.action[i].disabled = __H.$stopit
				break;
			}
		}
		return true
	}
	catch(e){
		__HLog.error(e,this);
		return false;
	}
}

function mba_common_stop(bStop,ev){
	try{
		if(!__HMBA || !__HMBA.isactive) return;

		if(mba_common_stopit(bStop,ev)){
			for(var i = ID_TIMEOUT1.length-1; i >= 0; i--) __HSys.timeStop(ID_TIMEOUT1[i])
			for(var i = ID_TIMEOUT1.length-1; i >= 0; i--) __HSys.timeStop(ID_TIMEOUT2[i])
			__HSys.timeStop(ID_TIMEOUT_MAIN);

			__H.$resource = oWno.ComputerName
		}
		return true;
	}
	catch(e){
		__HLog.error(e,this);
		return false;
	}
}

function mba_common_wait(sOpt,sOpt1,sOpt2,bOpt3){
	try{
		sOpt1 = typeof(sOpt1) == "string" ? sOpt1 : __HMBA.mod.divname
		if(sOpt == "start"){
			if(__HMBA.isactive) return;
			__HMBA.isactive = true
			mba_common_stop(false)
			ID_TIMEOUT1.push(window.setTimeout(function(){__MCommon.lockform(oFormAction,true)},0));
			if(sOpt1) ID_TIMEOUT1.push(window.setTimeout("__MCommon.lockform('" + sOpt1 + "',true)",0));

			if(typeof(sOpt2) == "string") __HLog.info(sOpt2,__HMBA.txt.info_css,true);
		}
		else if(sOpt == "stop" || sOpt == "init"){
			__HSys.timeStop(ID_TIMEINT_WAIT);

			if(sOpt1) ID_TIMEOUT1.push(window.setTimeout("__MCommon.lockform('" + sOpt1 + "',false)",0));
			if(sOpt == "init") mba_common_event(sOpt);
			ID_TIMEOUT1.push(window.setTimeout(function(){__MCommon.lockform(oFormAction,false)},0));
			__HMBA.isactive = false
		}
		else if(sOpt == "master"){
			if(sOpt1) document.body.style.cursor = oDivBody.style.cursor = sOpt2, __HMBA.master = true
			else __HMBA.master = false
		}
	}
	catch(e){
		__HLog.error(e,this);
		return false;
	}
	finally{
		/*
		var aOTables = document.getElementsByTagName("table")
		for(var i = 0,len = aOTables.length; i < len; i++){
			aOTables[i].refresh()
		}
		document.recalc(false)
		//__HUtil.kill(aOTables)
		*/
	}
}
