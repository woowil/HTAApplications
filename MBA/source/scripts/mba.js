// nOsliw HUI - HTML/HTA Application Framework Library (https://github.com/woowil/HTAFrameworks)
// Copyright (c) 2003-2013, nOsliw Solutions,  All Rights Reserved.
//**Start Encode**

var ID_TIMEOUT1			= [] // Used for invoking functions like threads (non-persistent)
var ID_TIMEOUT2			= [] // Used for invoking functions like threads (persistent)
var ID_TIMEOUT3			= [] // Used for invoking functions like threads (non-persistent). Remove on actions
var ID_TIMEOUT_MAIN		= null
var ID_TIMEINT1			= []
var ID_TIMEINT2			= []
var ID_TIMEINT_WAIT		= null
var ID_TIMEINT_PROGRESS	= null
var ID_TIMEINT_CHECK	= null
var ID_TIMEINT_CLOCK	= null
var ID_TIMEINT_DUMMY	= null
var ID_TIMEINT_PROCESS	= null

//////////////////////////////

__H.register(__H.UI.Window.HTA,"MBA","nOsliw MBA",function MBA(){
	this.date = new Date()
	this.buf  = new __H.Common.StringBuffer()

	this.initialize = function initialize(){
		
		for(var o in __H.$config){
			if(__H.$config.hasOwnProperty(o)){
				this["con_" + o] = __H.$config[o]
			}
		}

		this.sdate			= this.date.formatDateSV()
		this.sdate2			= this.date.formatYYYYMMDD() // 20050517
		this.sdate3			= this.date.formatYYYYMM()
		this.islocked		= true
		this.isactive		= this.isloaded = false
		this.title			= oMBA.applicationName
		this.userdomain		= oWno.UserDomain + "\\" + oWno.UserName
		this.user			= oWno.UserName
		this.ouser			= null
		this.comp			= oWno.ComputerName
		this.domain			= oWno.UserDomain
		this.issmallscreen	= window.screen.height < 800 ? true : false
		this.CharCollapse	= "&#x31;"
		this.CharExpand		= "&#x37;"

		this.registry		= null
		this.path_reg_base  = this.con_PathHKCU
		this.path_reg_hkcu  = this.con_PathHKCU + oMBA.version + "\\"
		this.path_reg_hklm  = this.con_PathHKLM + oMBA.version + "\\"
		this.localhost		= null
		this.appdata		= __HIO.appdata + this.con_PathAppData

		this.mnu			= [] // MENUBAR
		this.mnu_name		= ["Home","File","Edit","View","Tools","Modules","Help"]

		this.mdl			= []
		this.mdl_cur		= null
		this.mdl_dic1		= new ActiveXObject("Scripting.Dictionary")
		this.mdl_dic2		= new ActiveXObject("Scripting.Dictionary")
		this.mdl_dic3		= new ActiveXObject("Scripting.Dictionary")

		var objInf = function(bActive,sText,bDefault){
			this.isActive = bActive ? true : false
			this.isDefault = bDefault ? true : false
			this.text = typeof(sText) == "string" ? sText : ""
			this.text_default = this.isDefault ? this.text + " (default)" : this.text
			this.index = this.cur_inf_idx++
		}
		this.cur_lnf_idx	= 0
		this.lnf 			= [
			new objInf(true,"Menu",true,0),
			new objInf(false,"List",false,1),
			new objInf(false,"Tree",false,2)
		]
		this.cur_lnf		= this.lnf[0]

		var objMode = function(bActive,sText,bDefault){
			this.isActive = bActive ? true : false
			this.isDefault = bDefault ? true : false
			this.text = typeof(sText) == "string" ? sText : ""
			this.text_default = this.isDefault ? this.text + " (default)" : this.text
			this.index = this.cur_mode_idx++
		}
		this.cur_mode_idx	= 0
		this.mode			= [
			new objMode(true,"Full Control",true),
			new objMode(false,"Read-Only",false),
			new objMode(false,"Special Rights",false)
		]
		this.cur_mode		= this.mode[0]

		var objColor = function(bActive,sText,bDefault,sColNormal,sColDark){
			this.isActive = bActive ? true : false
			this.isDefault = bDefault ? true : false
			this.text = (sText || "")
			this.text_default = this.isDefault ? this.text + " (default)" : this.text
			this.col_normal = typeof(sColNormal) == "string" ? sColNormal : "#EEEEEE"
			this.col_dark = typeof(sColDark) == "string" ? sColDark : "#545454"
			this.index = this.cur_color_idx++
		}
		this.cur_color_idx = 0
		this.color			= [
			// http://html-color-codes.com/__H
			new objColor(true,"#CCCC99 | #336600",true,'#CCCC99','#336600'),
			new objColor(false,"#C1D2EE | #316AC5",false,'#C1D2EE','#316AC5'),
			new objColor(false,"#CCCCCC | #8A867A",false,'#CCCCCC','#8A867A'),
			new objColor(false,"#666600 | #999900",false,'#999900','#666600'),
			new objColor(false,"#0066CC | #99CCFF",false,'#99CCFF','#0066CC'),
			new objColor(false,"#336600 | #669933",false,'#336600','#669933'),
			new objColor(false,"DarkSeaGreen | Teal",false,'DarkSeaGreen','Teal'),
			new objColor(false,"Dark Blue | Blue",false,'Blue','DarkBlue'),
			new objColor(false,"#FF3300 | #FF9999",false,'#FF9999','#FF3300'),
			new objColor(false,"Silver | #EEEEEE",false,'#EEEEEE','Silver'),
			new objColor(false,"#000000 | #545454",false,'#545454','#000000')
		]
		this.cur_color		= this.color[0]

		var objSkins			= function(bActive,sText,bDefault,sCSS,sColStatus,iColIndex){
			this.isActive = bActive ? true : false
			this.isDefault = bDefault ? true : false
			this.text = typeof(sText) == "string" ? sText : ""
			this.text_default = this.isDefault ? this.text + " (default)" : this.text
			this.col_statusbar = typeof(sColStatus) == "string" ? sColStatus : "#ECE9D8"
			this.col_index = typeof(iColIndex) == "number" ? iColIndex : 0
			this.css_href = typeof(sCSS) == "string" ? sCSS : null
			this.index = this.cur_skins_idx++
		}
		this.cur_skins_idx	= 0
		this.skins			= [
		new objSkins(true,"Office XP",true,this.con_OfficeXP_CSS,"#ECE9D8",0),
		new objSkins(false,"Windows Classic",false,this.con_Classic_CSS,"#CCCCCC",5)
		]
		this.cur_skins		= this.skins[0]

		this.mod			= {
			author				: "Uuub~utrn Ukjquh".encode(),
			purpose				: "Ug5mt6 w2k65 ~LI Ihhtq42lqgv".encode(),
			credits				: "Uuub~utrn Ukjquh".encode() + "\nJavaScript Prototypes\nNotepad++",
			vendor				: "hMqjk~ Yujorkuhq".encode(),
			checked				: "atggt_okbufeeno".encode(),
			div					: oDivHome,
			list				: null, // my linked list for navigation
			time_start			: this.date,
			log_maxwait			: 15000,  // Used for log synchronization max wait
			log_res_isactive	: false, // Used for log synchronization
			log_err_isactive	: false, // Used for log synchronization
			time_lasttouched	: this.date, // Used for time interval..
			time_lockedout		: 60, // 60 minutes
			errfirsttime		: true,
			nav_go				: [],
			nav_go_idx			: 0,
			firsttime_docs		: true,
			buttom_wd1			: "80px",
			buttom_wd2			: "100px",
			buttom_wd3			: "120px",
			buttom_wd4			: "140px",
			buttom_wd5			: "220px",
			log_length			: 75,
			showIdle			: true,
			action_len			: 6,
			fadedelay1			: 250,
			fadedelay2			: 1000,
			menubar				: null
		}

		this.pth		= {
			cur				: oFso.GetAbsolutePathName("."), //oWsh.CurrentDirectory
			abt				: oFso.GetAbsolutePathName(this.con_PathAbt),
			bin				: oFso.GetAbsolutePathName(this.con_PathBin),
			htm				: oFso.GetAbsolutePathName(this.con_PathHtm),
			pic				: oFso.GetAbsolutePathName(this.con_PathPic),
			pic_url			: ((this.con_PathPic).replace(/[.]{0,1}\\/g,"/")).substr(1),
			pic_icon		: oFso.GetAbsolutePathName(this.con_PathPicIcon),
			pic_icon_url	: (this.con_PathPicIcon).replace(/\\/g,"/"),
			doc				: oFso.GetAbsolutePathName(this.con_PathDoc),
			css				: oFso.GetAbsolutePathName(this.con_PathCSS),
			mdl				: oFso.GetAbsolutePathName(this.con_PathModule),
			mdl_name		: (this.con_PathModule).replace(/^[.\\]*(.+)[\\]*$/g,"$1"),
			tmp				: this.appdata,
			tmp_unc			: "\\\\" + this.comp + "\\C$\\" + (this.appdata).substring(3,(this.appdata).length)
		}

		this.fls		= {
			autoitx				: this.pth.bin + "\\autoitx3.dll",
			regobjx				: this.pth.bin + "\\regobj.dll",
			adssecurityx		: this.pth.bin + "\\adssecurity.dll",
			setowner			: this.pth.bin + "\\setowner.exe",
			setacl				: this.pth.bin + "\\setacl.exe",
			netsh				: this.pth.bin + "\\netsh.exe",
			psexec				: this.pth.bin + "\\psexec.exe",
			psshutdown			: this.pth.bin + "\\psshutdown.exe",
			tsprof				: this.pth.bin + "\\tsprof.exe",
			fileacl				: this.pth.bin + "\\fileacl.exe",
			wbemtest			: __HIO.system32 + "\\wbem\\wbemtest.exe",

			mba_module			: this.pth.mdl + "\\" + this.con_FileModule,
			mba_program			: this.pth.cur + "\\" + oFso.GetFileName((document.URL).replace(/#/g,"")),
			mba_releasenotes	: this.pth.abt + "\\MBA Release Notes.txt",
			mba_license			: this.pth.abt + "\\MBA Licence.txt",
			mba_disclaimer		: this.pth.abt + "\\MBA Disclaimer.txt",
			mba_readme			: this.pth.abt + "\\MBA Readme.txt",
			mba_copyright		: this.pth.abt + "\\MBA Copyright.txt",

			divhead				: this.pth.htm + "\\mba-head.htm",
			divfoot				: this.pth.htm + "\\mba-foot.htm",
			divinfo				: this.pth.htm + "\\mba-info.htm",

			log_extra			: false,
			isLogResult			: true,

			log_whoami			: this.pth.cur + "\\logs\\" + this.sdate3 + "\\mba-whoami.log",
			log_result			: this.pth.cur + "\\logs\\" + this.sdate3 + "\\mba-" + this.sdate2 + "-" + this.user + "-" + this.comp + "-log-result.log",
			log_error			: this.pth.cur + "\\logs\\" + this.sdate3 + "\\mba-" + this.sdate2 + "-" + this.user + "-" + this.comp + "-log-error.log",

			/*
			"mig_robolog_users" : this.migration + "\\" + this.sdate3 + "\\mba-" + this.sdate2 + "-" + this.user + "-" + this.comp + "-robolog-users.log";
			"mig_robolog_groups" : this.migration + "\\" + this.sdate3 + "\\mba-" + this.sdate2 + "-" + this.user + "-" + this.comp + "-robolog-groups.log";
			"mig_robolog_apps_fdp" : this.migration + "\\" + this.sdate3 + "\\mba-" + this.sdate2 + "-" + this.user + "-" + this.comp + "-robolog-apps_fdp.log";
			"mig_robolog_apps_ks10" : this.migration + "\\" + this.sdate3 + "\\mba-" + this.sdate2 + "-" + this.user + "-" + this.comp + "-robolog-apps_ks10.log";
			"mig_robolog_mig" : this.migration + "\\" + this.sdate3 + "\\mba-" + this.sdate2 + "-" + this.user + "-" + this.comp + "-robolog-migration.log";
			"mig_robolog_log" : this.migration + "\\" + this.sdate3 + "\\mba-" + this.sdate2 + "-" + this.user + "-" + this.comp + "-robolog.log";
			"mig_robolog_takeowner" : this.migration + "\\" + this.sdate3 + "\\mba-" + this.sdate2 + "-" + this.user + "-" + this.comp + "-migration-takeowner.cmd";
			"mig_robolog_move" : this.migration + "\\" + this.sdate3 + "\\mba-" + this.sdate2 + "-" + this.user + "-" + this.comp + "-migration-move.cmd";
			"mig_diruse_src" : this.migration + "\\" + this.sdate3 + "\\mba-" + this.sdate2 + "-" + this.user + "-" + this.comp + "-migration-diruse-src.log";
			"mig_diruse_dst" : this.migration + "\\" + this.sdate3 + "\\mba-" + this.sdate2 + "-" + this.user + "-" + this.comp + "-migration-diruse-dst.log";
			"mig_robolog_extra" : this.migration + "\\" + this.sdate3 + "\\mba-" + this.sdate2 + "-" + this.user + "-" + this.comp + "-migration-result.log";
			"mig_copyfolders" : this.migration + "\\" + this.sdate3 + "\\mba-" + this.sdate2 + "-" + this.user + "-" + this.comp + "-robocopy-folders.ini";
			*/

			docs_dialog_help		: this.pth.htm + "\\mba-dialog-help.htm",
			docs_dialog_about		: this.pth.htm + "\\mba-dialog-about.htm",
			docs_dialog_busy		: this.pth.htm + "\\mba-dialog-busy.htm",
			docs_dialog_debug		: this.pth.htm + "\\mba-dialog-debug.htm",

			css_common				: this.pth.css + "\\mba.css", // oFso.GetAbsolutePathName(this.con_FileCSSHTA)
			css_common_url			: "file:///" + (this.pth.css + "\\mba.css").replace(/[.]{0,1}\\/g,"/"),
			docs_ini				: this.pth.doc + "\\mba-docs.ini",
			docs_workshop_url		: ((this.pth.doc).replace(/[.]{0,1}\\/g,"/")).substr(1) + "/mba-workshop.htm",

			div_text				: this.pth.tmp + "\\mba-" + this.sdate2 + "-" + this.user + "-saveas.txt",
			div_html				: this.pth.tmp + "\\mba-" + this.sdate2 + "-" + this.user + "-saveas.html"
		}

		this.wmi		= {}

		this.screen = {
			mode_3	: false,
			mode_4	: false,
			width	: oMBA.width,
			height	: window.screen.height*0.9,
			width2	: window.screen.width-50,
			height2	: window.screen.height*0.9,
			width3	: 750,
			height3	: 330,
			xpos	: window.screenLeft,
			ypos	: window.screenTop
		}

		this.txt = {
			line			: "------------------------------------------------------------------------------\n",
			filter1			: "HTA Files (*.txt;*.log;*.vbs;*.js;*.pl;*.wsf)\0*.txt;*.log;*.vbs;*.js;*.pl;*.wsf\0" + "All Files (*.*)\0*.*\0\0",
			filter2			: "Script Files (*.vbs;*.js;*.pl;*.wsf)\0*.vbs;*.js;*.pl;*.wsf\0\0",
			locked			: "\u25A1\u25A1\u25A1 &nbsp; H T A &nbsp;&nbsp; I S &nbsp;&nbsp; L O C K E D . &nbsp;&nbsp;&nbsp; U N L O C K &nbsp;&nbsp; I T ! &nbsp; \u25A1\u25A1\u25A1",
			unlocked		: "\u25A1\u25A1\u25A1 &nbsp; H T A &nbsp;&nbsp; I S &nbsp;&nbsp; N O W &nbsp;&nbsp; U N L O C K E D ! &nbsp; \u25A1\u25A1\u25A1",
			notimpl			: "\u25A1\u25A1\u25A1 &nbsp; N O T &nbsp;&nbsp; Y E T &nbsp;&nbsp; I M P L E M E N T E D ! &nbsp; \u25A1\u25A1\u25A1",
			deactivated		: "\u25A1\u25A1\u25A1 &nbsp; T H I S &nbsp;&nbsp; S E C T I O N &nbsp;&nbsp; I S &nbsp;&nbsp; D E A C T I V A T E D ! &nbsp; \u25A1\u25A1\u25A1",
			nocando			: "Avn2tq5 mk6jv2u6 gj h2kkBgj5".encode(),
			credentials		: "c r e d e n t i a l s &nbsp; u p d a t e d !",
			mba_info		: "\u25A1\u25A1 " + this.mod.purpose + ' \u25A1\u25A1 &nbsp;&nbsp; <span class="mba-head-3"> Copyright &copy; \u25A1\u25A1\u25A1\u25A1 ' + this.mod.vendor + ' \u25A1\u25A1\u25A1\u25A1 version: ' + oMBA.version,
			mba_info2		: '<span class="mba-head-3" style="color:#545454;">' + this.mod.vendor + ' :: Copyright &copy; 2003-2008 :: Module Based HTA Application (MBA)</span>'
		}
	}

	this.onLoad = function onLoad(){
		try{
			
			document.title = "..initializing 1(3)"
			oWsh.CurrentDirectory = oFso.GetAbsolutePathName(".")
			
			// DO NOT MOVE. MUST BE LOADED BEFORE ANYTHING ELSE
			this.initialize()
			
			__HLog.log		= __MLog.log
			__HLog.info		= __MLog.info
			__HLog.logInfo	= __MLog.logInfo
			__HLog.logPopup	= __MLog.logPopup
			__HLog.error	= __MLog.error

			document.body.onkeydown = new Function("mba_common_event()")
			document.body.onhelp = new Function("oFormAction.help.onclick()")
			
			// DO NOT MOVE! MUST BE AFTER hui declaration
			__HFolder.create("C:\\Temp",this.pth.tmp,this.pth.cur + "\\logs\\" + this.sdate3)

			mba_common_menu()
			if(this.con_loadmodule == 1) ID_TIMEOUT1.push(window.setTimeout(function(){__MModule.initialize(true)},0))
	
			document.title = "..initializing 2(3)"

			this.rabbit()
			mba_common_formbehavior("behavior")

			oFormMenu[this.mod.checked].value = "".pass(1)
			this.buf.empty()

			this.buf.append('\n<table width="100%" align="center" class="mba-table-1" border="1" cellspacing="2" cellpadding="0">')
			this.buf.append('\n<tr><td>User</td><td>' + this.user + '</td></tr>')
			this.buf.append('\n<tr><td>Domain</td><td>' + this.domain + '</td></tr>')
			this.buf.append('\n<tr><td>Computer</td><td>' + this.comp + '</td></tr>')
			this.buf.append('\n<tr><td>Date</td><td>' + this.sdate + '</td></tr>')
			this.buf.append('\n<tr><td>Small Screen</td><td>' + this.issmallscreen + '</td></tr>')
			this.buf.append('\n<tr><td>Version</td><td>' + oMBA.version + '</td></tr>')
			this.buf.append('\n</table>')
			oDivHomeSesSub.innerHTML = this.buf.toString()
			this.buf.empty()

			oFormOptions.options_hideresult.onclick = new Function("mba_common_formactivate(this);")
			oFormOptions.options_hideresult.checked = window.screen.height <= 768 ? true : false;
			oFormOptions.options_loglength.onchange = new Function("__HMBA.mod.log_length=this.value");
			oFormOptions.options_loglength.selectedIndex = 0
			oFormOptions.options_loglength.onchange()
	
			oFormDebug.dbg_style_text.value = this.fls.css_common
			
			document.title = "..initializing 3(3)"
			this.rabbit()
			//__HHTA.position(this.screen.width3,this.screen.height3)
		
			window.setTimeout(function(){__MLog.whoami()},0)
			
			ID_TIMEINT_CHECK = setInterval(function(){__MCommon.intervalCheck()},60000);
			ID_TIMEINT_CLOCK = setInterval(function(){__MCommon.intervalClock()},2000);
		}
		catch(ee){
			__HLog.popup("__HMBA.onLoad(): " + ee.description)
			__HLog.error(ee,this)
		}
		finally{
			document.title = oMBA.applicationName
			document.recalc(true)
			this.isloaded = true
		}
	}

	this.onunLoad = function onunLoad(){
		try{
			
		}
		catch(ee){}

		for(var i = 0; i < ID_TIMEOUT3.length; i++) __HSys.timeStop(ID_TIMEOUT3[i])
		for(var i = 0; i < ID_TIMEINT1.length; i++) __HSys.timeStop(ID_TIMEINT1[i])

	}

	this.refresh = function refresh(){
		var t = document.title
		self.focus()
		document.title = "..refreshing"
		__HAuto.setTitle(document.title)
		__HAuto.refresh()
		oTblTop.refresh()
		document.recalc(true)
		document.title = t
		if(this.isloaded) self.focus()
	}

	this.rabbit = function rabbit(){
		try{
			if(!window.oDivLogin){
				var s3 = "Ent"+"er key to unl"+"ock"
				var o = __H.byClone("div")
				o.id = "oDivLogin"
				o.innerHTML = '\n<FORM name="oFormLogin">' +
				'\n<TABLE border=0 height="100%" width="100%">' +
				'\n<TR>' +
				'\n\t<TD vAlign="top" align="middle"><BR><BR>' +
				'\n\t<TABLE cellSpacing=0 cellPadding=0 width=650 border=0>' +
				'\n\t<TR>' +
				'\n\t\t<TD>' +
				'\n\t\t<DIV class="mba-div-sup" id="oDivRabbitSup" title="Lo'+'gon">Lo'+'gon</DIV>' +
				'\n\t\t</TD>' +
				'\n\t</TR>' +
				'\n\t<TR>' +
				'\n\t\t<TD vAlign="top" align="right">' +
				'\n\t\t<SPAN class="mba-span-smal"><A onclick="__MDoc.dialogAbout()" href="javascript:void(0)">Ab' + 'out</A> :: <A onclick="location.reload()" href="javascript:void(0)">Rel' + 'oad</A> :: <A onclick="__HHTA.closeHTA()" href="javascript:void(0)">E' + 'xit</A> &nbsp;</SPAN><BR><BR>' +
				'\n\t\t</TD>' +
				'\n\t</TR>' +
				'\n\t<TR>' +
				'\n\t\t<TD align="middle">' +
				'\n\t\t<DIV class="mba-div-sub" id="oDivRabbitSub">' +
				'\n\t\t<FIELDSET class="mba-fieldset-2">' +
				'\n\t\t<LEGEND id="oLegRabbit"><INPUT title="' + s3 + '" type=password maxLength=20 size=15 value="" name=catc'+'hmeifyo'+'ucan></LEGEND>' +
				'\n\t\t<MARQUEE id="oMarRabbit" onmouseover="this.stop()" onmouseout="this.start()" scrollAmount=1 hspace=4 vspace=8 scrollDelay=80 direction="up" loop="infinite" width="100%" height=100 align="left"></MARQUEE>' +
				'\n\t\t</FIELDSET>' +
				'\n\t\t</DIV><BR>' +
				'\n\t\t</TD>' +
				'\n\t</TR>' +
				'\n\t</TABLE>' +
				'\n\t</TD>' +
				'\n</TR>' +
				'\n<TR>' +
				'\n\t<TD vAlign="bottom" align="middle" height=30>' +
				'\n\t<DIV class="mba-div-filter" id="oDivRabbitText">' + this.txt["mba_i"+"nfo2"] + '<BR><BR>&nbsp;' +
				'\n\t</DIV>' +
				'\n\t</TD>' +
				'\n</TR>' +
				'\n</TABLE>' +
				'\n</FORM>'

				document.body.appendChild(o)
				window.setTimeout(function(){oMarRabbit.innerText=__HFile.readall(__HMBA.fls.mba_releasenotes)},0)
			}
			else{
				if(oWsh.AppActivate("nOsliw :: MBA :: Splash")){
					oWsh.SendKeys("%{F4}");
				}

				__HElem.hide(oTblTop,oDivLogin,oDivRabbitText)
				if(__HElem.isHidden(oDivRabbitSub)) oDivRabbitSup.onmousedown()				
				__HElem.show(oDivLogin)
				
				__HHTA.position(this.screen.width3,this.screen.height3)
				window.setTimeout(function(){__HElem.transition(oDivRabbitText)},this.mod.fadedelay2)
				
				__HUI.b_fullscreen = false
				this.islocked = true
			}
		}
		catch(ee){
			__HLog.error(ee,this)
			return false;
		}
	}
})

var __HMBA = __M = new __H.UI.Window.HTA.MBA()
