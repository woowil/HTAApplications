// Copyright© 2003-2008. nOsliw Solutions. All rights reserved.
//**Start Encode**

function management_wmi(sOpt){
	try{
		var oForm = oFormManagementWmi, aClass
		if(sOpt == "connect"){
			if(!__HWMI.setServiceWMI(__HMBA.wmi.u_server,__HMBA.wmi.u_namespace,__HMBA.wmi.u_user,__HMBA.wmi.u_pass)) return false
			return true
		}
		else if(sOpt == "method"){
			__HLog.logInfo("# Getting WMI class method objects using class " + __HMBA.wmi.u_class,null,true)
			if(!(aOClass = __HWMI.getMethods(__HMBA.wmi.u_class))){
				__HLog.popup("## Unable to get WMI method on class " + __HMBA.wmi.u_class);
				return false;
			}
			__HLog.log("## Creating table for methods using ")
			var oBody = __H.byIds("oBdyManagementWmiMet"), oRow, oCell
			var oCap = __H.byIds("oCapManagementWmiMet")
			oCap.innerHTML = "<i>Methods for</i> <b>" + __HMBA.wmi.u_class + "</b>"
			for(var i = oBody.rows.length; i > 0; i--) oBody.deleteRow(i-1);
			
			// Its faster to clone node
			// http://archive.devwebpro.com/devwebpro-39-20030514OptimizingJavaScriptforExecutionSpeed.html
			for(var i = 0, n, iLen = aOClass.length; i < iLen; i++){
				var o = aOClass[i], oRow
				oRow = oBody.insertRow(oBody.rows.length), oRow.style.verticalAlign = 'top', n = (oBody.rows.length).toNumberZero();
				oCell = oRow.insertCell(), oCell.innerText = n, oCell.width = 22, oCell.align = "center"
				oCell = oRow.insertCell(), oCell.innerText = o.name + "()", oCell.className = "mba-table-1TD", oCell.noWrap = true, oCell.width = "1%"
				oCell = oRow.insertCell(), oCell.noWrap = true
				oCell.innerText = "IN PARAMETER\n" + o.inparam + "\n\nOUT PARAMETER\n" + o.outparam
				oCell = oRow.insertCell(), oCell.innerText = o.qualifiers + " "
				oCell = oRow.insertCell(), oCell.innerText = o.description + " "
				oCell = oRow.insertCell(), oCell.innerText = n, oCell.width = 22, oCell.align = "center"
			}
			__HPanel.activate("oDivManagementWmiPanel",2);
			return true
		}
		else if(management_wmi("connect")){
			__HLog.logInfo("# Getting WMI class property objects using class " + __HMBA.wmi.u_class,null,true)
			var sSimple = oForm.wmi_opt_getdescription.checked ? null : "simple"
			var oBody = __H.byIds("oBdyManagementWmiPro"), oRow, oCell
			var oCap = __H.byIds("oCapManagementWmiPro")

			oCap.innerHTML = "<i>Properties for</i> <b>" + __HMBA.wmi.u_class + "</b>"
			for(var i = oBody.rows.length; i > 0; i--) oBody.deleteRow(i-1);
			//__HUtil.sleep(50)
			if(!(aClass = management_wmi_getclass(sSimple,__HMBA.wmi.u_class,oForm.query.value,"\n",oBody))){
				__HLog.popup("## Unable to get WMI property on class " + __HMBA.wmi.u_class);
				return false;
			}

			__HLog.log("## Creating table for properties using query: " + oForm.query.value)
			for(var i = n = 0, r, iLen = aClass.length; i < iLen; i++, n++){
				for(var o in aClass[i]){
					if(!aClass.hasOwnProperty(o)) continue
					var oo = aClass[i][o], d = oo.description + " ", v = oo.value + " "
					oRow = oBody.insertRow(n), oRow.style.verticalAlign = 'top', r = (n+1).toNumberZero();
					oCell = oRow.insertCell(), oCell.innerText = n, oCell.width = 22, oCell.align = "center"
					oCell = oRow.insertCell(), oCell.innerText = o, oCell.className = "mba-table-1TD", oCell.noWrap = true, oCell.width = "1%"
					oCell = oRow.insertCell(), oCell.innerText = v
					oCell = oRow.insertCell(), oCell.innerText = d
					oCell = oRow.insertCell(), oCell.innerText = r, oCell.width = 22, oCell.align = "center"
				}
				if(aClass.length > (1+i)){
					oRow = oBody.insertRow(n+++1)
					oCell = oRow.insertCell(), oCell.innerText = " "
					oCell.className = "mba-head-2", oCell.colSpan = 5
				}
			}
			oFormAction.action[0].disabled = oFormAction.action[1].disabled = false
			if(oForm.wmi_opt_getmethods.checked) oFormAction.action[1].onclick()

			var Dummy1 = setTimeout("management_wmi_getscript()",3);
			return true
		}
		return false
	}
	catch(e){
		__HLog.error(e,this);
		return false;
	}
	finally{
		//__HUtil.kill(aClass)
	}
}

function management_wmi_getscript(sLanguage){
	try{
		if(__HMBA.wmi.u_class){
			sLanguage = sLanguage ? sLanguage : oFormManagementWmi.wmi_scr_language.value
			__HLog.logInfo("### Getting WMI script on class: " + __HMBA.wmi.u_class + " for scripting language: " + sLanguage)
			document.all["oTextareaWmiScr"].innerText = ""
			var sStream
			if(sStream = management_wmi_language(__HMBA.wmi.u_class,sLanguage,__HMBA.wmi.u_server,__HMBA.wmi.u_namespace,__HMBA.wmi.u_user,__HMBA.wmi.u_pass)){
				document.all["oTextareaWmiScr"].innerText = sStream
			}
		}
	}
	catch(e){
		__HLog.error(e,this);
		return false;
	}
}

function management_wmi_getclass(sOpt,sClass,sQuery,sBreak,oTBody){
	var aProperties, sGetObject
	try{
		sQuery = !__H.isStringEmpty(sQuery) ? sQuery : "Select * from " + sClass
		sBreak = !__H.isStringEmpty(sBreak) ? sBreak : "\n  "
		var oService = __HWMI.getServiceWMI();
		var oRe = new RegExp("Select[ \t]+(.+)[ \t]+from[ \t]+([a-z0-9_\-]+)[ \t]*(.*)","ig")
		var bSelect = false, oRe1

		if(sQuery.match(oRe)){
			var d_select = new ActiveXObject("Scripting.Dictionary");
			var a = (RegExp.$1).split(/[, \t]{1,}/ig)
			if(a[0] != "*"){
				for(var i = 0, iLen = a.length; i < iLen; i++) d_select.Add(a[i],a[i])
				if(sQuery.isSearch(/ where (.*)/ig)){
					var t = RegExp.$1
					t = t.replace(/[ \t]{2,}/g," ")
					t = t.replace(/[ \t]*=[ \t]*/g,"=")
					t = t.replace(/[ \t]+and[ \t]+/ig," ")
					t = t.replace(/[ \t]+or[ \t]+/ig," ")
					t = t.replace(/[<>]+/g,"")
					t = t.replace(/["']+/g,"")
					t = t.replace(/[ ]/g,"=")
					a = t.split("=")
					for(var i = 0, iLen = a.length; i < iLen; i += 2){
						if(!d_select.Exists(a[i])) d_select.Add(a[i],a[i])
					}
				}
				bSelect = true
			}
			__HUtil.kill(a)
		}
		else return false

		var bTBody = typeof(oTBody) == "object" ? true : false // Can directly populate a table using TBody
		var bSimple = sOpt == "simple" ? true : false
		var bShow = sOpt == "show" ? true : false

		sGetObject = bSimple ? "getobject_simple" : "getobject"
		var aProperties = __HWMI.getProperty(sClass)
		var ENUM_RUNTIME = 120
		if(aProperties){
			var startTime = new Date()
			var oColItems = oService.ExecQuery(sQuery,"WQL",48), r = 0, rr;
			if(bSimple || bShow) var aClass = []
			else if(bTBody){
				__HLog.log("## Creating table for properties using query: " + sQuery)
				var r = oTBody.rows.length, rr, oRow, oCell
				
			}
			var iLen = aProperties.length
			for(var i = 0, oEnum = new Enumerator(oColItems); !oEnum.atEnd() && !__H.$stopit; oEnum.moveNext(), i++){
				var oItem = oEnum.item();
				if(!bTBody){
					aClass[i] = [];
					if(bShow && i > 0) sClass = new __H.Common.StringBuffer(sClass + "\n")
				}
				else if(r > 0){ // do not create a row in the beginning
					oRow = oTBody.insertRow((r++))
					oCell = oRow.insertCell(), oCell.innerText = "instance - " + (i).toNumberZero(), oCell.className = "mba-head-2", oCell.colSpan = 5
				}

				for(var j = 0; j < iLen && !__H.$stopit; j++, r++){
					var n = aProperties[j].name, v, t
					if(bSelect && !d_select.Exists(n)) continue // Check if the query string has "*". If not, then exclude unchosen fields
					if(aProperties[j].isarray){
						//t = (v=management_wmi_vbarray(oItem[n],n,sBreak)) ? v.stream : " "
						t = __HWMI.wmiArray(oItem[n])
					}
					else if(aProperties[j].isdatetime) t = management_wmi_date(oItem[n])
					else t = oItem[n]
					if(bShow) sClass.append("\n" + n + ": " + t)
					else if(bTBody){
						oRow = oTBody.insertRow(r), oRow.style.verticalAlign = 'top', rr = (r+1).toNumberZero();
						oCell = oRow.insertCell(), oCell.innerText = rr, oCell.width = 22, oCell.align = "center"
						oCell = oRow.insertCell(), oCell.innerText = n, oCell.className = "mba-table-1TD", oCell.noWrap = true, oCell.width = "1%"
						oCell = oRow.insertCell(), oCell.innerText = t + " "
						oCell = oRow.insertCell()
						oCell.innerText = bSimple ? " " : aProperties[j].description
						oCell = oRow.insertCell(), oCell.innerText = rr, oCell.width = 22, oCell.align = "center"
					}
					else {
						aClass[i][n] = {
							name : n,
							value : t
						}
						if(!bSimple){
							aClass[i][n].description = aProperties[j].description
							aClass[i][n].string = n + ": " + t
						}
					}				
				}
				if(bTBody){
					// continue if less than 1200 rows or less than 2 minutes
					var nowTime = new Date()
					var diffTime = nowTime.getDiff(startTime)
					if(r < 1200 || diffTime.getTotalSeconds() < ENUM_RUNTIME) continue 
					
					var perSecond = i/diffTime.getTotalSeconds()
					var estimatedSeconds = Math.round(oEnum.getSize()/perSecond)
					var endTime = new Date()
					endTime.setTime(diffTime.getTime()+(estimatedSeconds*1000))
					
					var s = "The WQL query enumeration for '" + sQuery + "' is taking a lot of time and consuming CPU.\n" + 
					"\n* " + i + " instances have been enumerated (out of " + oEnum.getSize() + ") after " + diffTime.getTotalSeconds() + " seconds." +
					"\n* Estimated query time is '" + estimatedSeconds + "' seconds (" + endTime.formatHHhMMmSSs() + ")" +
					"\n\nStop Enumeration?"
					if(oWsh.Popup(s,120,"WMI Enumeration",32 + 5) == 6) break; // Yes
					
					ENUM_RUNTIME += (diffTime.getTotalSeconds()+120)
				}
			}
		}
		return (bTBody ? (r > 0) : aClass)
	}
	catch(e){
		__HLog.error(e,this);
		return false;
	}
	finally{
		if(bSelect) d_select.RemoveAll(), delete d_select
		if(bShow) __HLog.log(sClass.toString())
		__HUtil.kill(aProperties)
	}
}

function management_wmi_date(d){
	try{
		if(d != null){
			d = d.replace(/[\*]/g,0);
			return d.substring(0,4) + "-" + d.substring(4,6) + "-" + d.substring(6,8) + " " + d.substring(8,10) + ":" + d.substring(10,12) + ":" + d.substring(12,14) + "." + d.substring(15,18);
		}
		return false;
	}
	catch(e){
		return false;
	}
}

function management_wmi_vbarray(aVBarray,sVBName,sBreak){
	try{
		if(aVBarray != null){ // VBArray object
			var aVB = new VBArray(aVBarray).toArray();
			var a = []
			sBreak = sBreak ? sBreak : "\n  "
			for(var j = 0, sVB = "", iLen = aVB.length; j < iLen; j++){
				sVB = sVB.concat((j == 0 ? "" : sBreak) + aVB[j])
				a.push(aVB[j])
			}
			aVB.stream = (j == 1) ? sVB.replace(/.+-[0-9]{1,2}:  (.+)$/ig,"$1") : sVB;
			aVB.array = a;
			return aVB;
		}
		return false;
	}
	catch(e){
		return false;
	}
}

function management_wmi_language(sClass,sLanguage,sComputer,sNameSpace,sUser,sPass,sImpersonation,sLocale,sAuthentication,sAuthority,iPrivilege,iSecurityFlags){
	try{
		var sHtml = new __H.Common.StringBuffer()
		if(!__HWMI.setServiceWMI(sComputer,sNameSpace,sUser,sPass)) return false
		var oClass = (__HWMI.getServiceWMI()).Get(sClass);
		
		sComputer = sComputer ? sComputer : oWno.ComputerName
		sUser = sUser ? sUser : null;
		sPass = sPass ? sPass : null;
		sNameSpace = sNameSpace ? sNameSpace : "root\\CIMV2"
		sLocale = sLocale ? sLocale : "MS_409"; //MS_409 is American English
		sImpersonation = sImpersonation ? sImpersonation : 3;
		sAuthentication = sAuthentication ? sAuthentication : 6; // Pkt(4): Authenticates that all data received is from the expected client.
		// Either NTLM {authority=ntlmdomain:DomainName} or Kerberos :{authority=kerberos:DomainName\ServerName}
		// If sUser=DOMAIN\username then sAuthority must be null
		sAuthority = sAuthority ? sAuthority : null;
		iPrivilege = iPrivilege ? iPrivilege : -1; // ???: All Privileges, 7: Security
		iSecurityFlags = iSecurityFlags ? iSecurityFlags : 128; // Connection timeout: 0 is indefinitely, 128 sec is max
		var bLocal = (!sComputer || sComputer == "." || sComputer.toUpperCase() == oWno.ComputerName.toUpperCase())
		
		if(sLanguage.toLowerCase() == "jscript"){ // JScript
			sHtml.append("// Made by: " + __HMBA.mod.purpose + "\n// Generated by: " + document.title);
					sHtml.append("\n// " + __HLang.disclaimer);
					sHtml.append("\n// http://msdn.microsoft.com/library/en-us/wmisdk/wmi/win32_classes.asp");
					var h = "\n\nWScript.Echo(\"[ WMI Class: \\\\" + sComputer + "\\" + sNameSpace + ":" + sClass + " ]\");\n\n";
					sHtml.append(h.replace(/\\/g,"\\\\"))
					if(bLocal){
						sHtml.append("wmi_" + sClass.toLowerCase() + "(\"" + sComputer + "\",\"" + sNameSpace.replace(/\\/g,"\\\\") + "\");");
						sHtml.append("\n\nfunction wmi_" + sClass.toLowerCase() + "(sComputer,sNameSpace,sImpersonation,sAuthentication){");
					}
					else{
						sHtml.append("wmi_" + sClass.toLowerCase() + "(\"" + sComputer + "\",\"" + sNameSpace.replace(/\\/g,"\\\\") + "\");");
						sHtml.append("\n\nfunction wmi_" + sClass.toLowerCase() + "(sComputer,sNameSpace,sUser,sPass,sImpersonation,sLocale,sAuthentication,sAuthority,iPrivilege,iSecurityFlags){");
					}

					sHtml.append("\n\ttry{\n\t\tvar sHtml, sHtmlAll = \"\";");
					sHtml.append("\n\t\tsComputer = sComputer ? sComputer : \".\";");
					sHtml.append("\n\t\tsNameSpace = sNameSpace ? sNameSpace : \"root\\\\CIMV2\";");

					if(bLocal){
						sHtml.append("\n\t\tsImpersonation = sImpersonation ? sImpersonation : \"Impersonate\";");
						sHtml.append("\n\t\tsAuthentication = sAuthentication ? sAuthentication : \"Connect\";");
						sHtml.append("\n\t\tvar oService = GetObject(\"winmgmts:{impersonationLevel=\" + sImpersonation + \",authenticationLevel=\" + sAuthentication + \",(Security)}!\\\\\\\\\" + sComputer + \"\\\\\" + sNameSpace);");
					}
					else {
						sHtml.append("\n\t\tsUser = sUser ? sUser : null;");
						sHtml.append("\n\t\tsPass = sPass ? sPass : null;");
						sHtml.append("\n\t\tsLocale = sLocale ? sLocale : \"MS_409\";");
						sHtml.append("\n\t\tsImpersonation = sImpersonation ? sImpersonation : 3;");
						sHtml.append("\n\t\tsAuthentication = sAuthentication ? sAuthentication : 4;");
						sHtml.append("\n\t\tsAuthority = sAuthority ? sAuthority : null;");
						sHtml.append("\n\t\tiPrivilege = iPrivilege ? iPrivilege : 7;");
						sHtml.append("\n\t\tiSecurityFlags = iSecurityFlags ? iSecurityFlags : 128;");

						sHtml.append("\n\n\t\tvar oLocator = new ActiveXObject(\"WbemScripting.SWbemLocator\");");
						sHtml.append("\n\t\tvar oService = oLocator.ConnectServer(sComputer,sNameSpace,sUser,sPass,sLocale,sAuthority,iSecurityFlags,null);");
						sHtml.append("\n\t\toService.Security_.ImpersonationLevel = sImpersonation;");
						sHtml.append("\n\t\toService.Security_.AuthenticationLevel = sAuthentication;");

						sHtml.append("\n\t\tif(iPrivilege == -1){");
						sHtml.append("\n\t\t\tfor(var p = 1; p < 27; p++){");
						sHtml.append("\n\t\t\t\ttry{");
						sHtml.append("\n\t\t\t\t\toService.Security_.Privileges.Add(p); // http://msdn.microsoft.com/library/default.asp?url=/library/en-us/wmisdk/wmi/wbemprivilegeenum.asp");
						sHtml.append("\n\t\t\t\t}");
						sHtml.append("\n\t\t\t\tcatch(ee){}");
						sHtml.append("\n\t\t\t}\n\t\t}");
						sHtml.append("\n\t\telse oService.Security_.Privileges.Add(7);");
					}
					sHtml.append("\n\t\tvar oColItems = oService.ExecQuery(\"Select * from " + sClass + "\",\"WQL\",48);");
					sHtml.append("\n\n\t\tfor(var oItem, i = 1, d, v, oEnum = new Enumerator(oColItems); !oEnum.atEnd(); oEnum.moveNext(), i++){");
					sHtml.append("\n\t\t\toItem = oEnum.item(), sHtml = \"\\n [INSTANCE-\" + i + \"]\";");

					for(var oEnum = new Enumerator(oClass.properties_); !oEnum.atEnd(); oEnum.moveNext()){
						var oItem = oEnum.item();
						if(oItem.IsArray){ // Only VBArrays
							/*
							sHtml.append("\n\t\t\tif(oItem." + oItem.Name + " != null){ // VBArray object");
							sHtml.append("\n\t\t\t\tvar a" + oItem.Name + " = new VBArray(oItem." + oItem.Name + ").toArray();");
							sHtml.append("\n\t\t\t\tfor(var j = 0; j < a" + oItem.Name + ".length; j++){");
							sHtml.append("\n\t\t\t\t\tsHtml += \"\\n  " + oItem.Name + "-\" + (j+1) + \":  \" + a" + oItem.Name + "[j];");
							sHtml.append("\n\t\t\t\t}");
							sHtml.append("\n\t\t\t}");
							sHtml.append("\n\t\t\telse sHtml += \"\\n  " + oItem.Name + ": undefined\";");
							*/
							sHtml.append("\n\t\t\tsHtml = sHtml + ((v=management_wmi_vbarray(oItem." + oItem.Name + ",'" + oItem.Name + "')) ? v.stream : \"\\n  " + oItem.Name + ": \" + \"undefined\");")
						}
						else if(oItem.CIMType == 101){ // If it is a DateTime object
							/*
							sHtml.append("\n\t\t\tif(oItem." + oItem.Name + " != null){");
							sHtml.append("\n\t\t\t\tvar d = oItem." + oItem.Name + ";");
							sHtml.append("\n\t\t\t\tvar sDateTime = d.substring(0,4) + \"-\" + d.substring(4,6) + \"-\" + d.substring(6,8) + \" \" + d.substring(8,10) + \":\" + d.substring(10,12) + \":\" + d.substring(12,14);");
							sHtml.append("\n\t\t\t\tsHtml += \"\\n  " + oItem.Name + ": \" + sDateTime;");
							sHtml.append("\n\t\t\t}");
							sHtml.append("\n\t\t\telse sHtml += \"\\n  " + oItem.Name + ": undefined\";");
							*/
							sHtml.append("\n\t\t\tsHtml += \"\\n  " + oItem.Name + ": \" + ((d=management_wmi_date(oItem." + oItem.Name + ")) ? d : \"undefined\");")
						}
						else sHtml.append("\n\t\t\tsHtml += \"\\n  " + oItem.Name + ": \" + oItem." + oItem.Name + ";")
					}

					sHtml.append("\n\t\t\tsHtmlAll += sHtml;")
					sHtml.append("\n\t\t\tWScript.Echo(sHtml);")
					sHtml.append("\n\t\t\tWScript.Sleep(1);")
					sHtml.append("\n\t\t}\n\t}")
					sHtml.append("\n\tcatch(e){")
					sHtml.append("\n\t\tWScript.Echo(\"Error:\\nNumber: \" + (e.number & 0xFFFF) + \"\\nDescription: \" + e.description + \" - Unable to connect to system. WMI may not be installed.\");")
					sHtml.append("\n\t\treturn false;")
					sHtml.append("\n\t}")
					sHtml.append("\n\tfinally{")
					sHtml.append("\n\t\treturn sHtmlAll;")
					sHtml.append("\n\t}")
					sHtml.append("\n}\n\n")

					sHtml.append(management_wmi_date + "\n\n" + management_wmi_vbarray)

					sHtml.append("\n\nWScript.Echo(\"\\nD O N E!\")\n")
				}
				else if(sLanguage.toLowerCase() == "vbscript"){ //VBscript
					sHtml.append("' Made by: " + __HMBA.mod.purpose + "\n' Generated by: " + document.title);
					sHtml.append("\n' " + __HLang.disclaimer);
					sHtml.append("\n' http://msdn.microsoft.com/library/en-us/wmisdk/wmi/win32_classes.asp");
					sHtml.append("\n\nWScript.Echo \"[ WMI Class: \\\\" + sComputer + "\\" + sNameSpace + ":" + sClass + " ]\"\n\n");
					if(bLocal){
						sHtml.append("wmi_" + sClass.toLowerCase() + " \"" + sComputer + "\",\"" + sNameSpace + "\",\"" + sImpersonation + "\",\"" + sAuthentication + "\"");
						sHtml.append("\n\nSub wmi_" + sClass.toLowerCase() + "(strComputer,strNameSpace,strImpersonation,strAuthentication)");
					}
					else{
						sHtml.append("wmi_" + sClass.toLowerCase() + " \"" + sComputer + "\",\"" + sNameSpace + "\"");
						sHtml.append("\n\nSub wmi_" + sClass.toLowerCase() + "(strComputer,strNameSpace,strUser,strPass,strImpersonation,strLocale,strAuthentication,strAuthority,intPrivilege,intSecurityFlags)");
					}

					sHtml.append("\n\tOn Error Resume Next\n\tErr.Clear");
					sHtml.append("\n\n\tstrHtml = \"\" : i = 0");
					sHtml.append("\n\tIf IsEmpty(strComputer) Or IsNull(strComputer) Then strComputer = \".\" End If");
					sHtml.append("\n\tIf IsEmpty(strNameSpace) Or IsNull(strNameSpace) Then strNameSpace = \"root\\CIMV2\" End If");

					if(bLocal){
						sHtml.append("\n\tIf IsEmpty(strImpersonation) Or IsNull(strImpersonation) Then strImpersonation = \"Impersonate\" End If");
						sHtml.append("\n\tIf IsEmpty(strAuthentication) Or IsNull(strAuthentication) Then strAuthentication = \"Connect\" End If");
						sHtml.append("\n\tSet objService = GetObject(\"winmgmts:{impersonationLevel=\" & strImpersonation & \",authenticationLevel=\" & strAuthentication & \",(Security)}!\\\\\" & strComputer & \"\\\" &  strNameSpace)");
					}
					else {
						sHtml.append("\n\tIf IsEmpty(strUser) Or IsNull(strUser) Then strUser = null End If");
						sHtml.append("\n\tIf IsEmpty(strPass) Or IsNull(strPass) Then strPass = null End If");
						sHtml.append("\n\tIf IsEmpty(strLocale) Or IsNull(strLocale) Then strLocale = \"MS_409\" End If");
						sHtml.append("\n\tIf Not IsNumeric(intImpersonation) Then intImpersonation = 3 End If");
						sHtml.append("\n\tIf Not IsNumeric(strAuthentication) Then strAuthentication = 4 End If");
						sHtml.append("\n\tIf IsEmpty(strAuthority) Or IsNull(strAuthority) Then strAuthority = Nothing End If");
						sHtml.append("\n\tIf Not IsNumeric(intSecurityFlags) Then intSecurityFlags = 128 End If");
						sHtml.append("\n\tIf Not IsNumeric(intPrivilege) Then intPrivilege = 7 End If");

						sHtml.append("\n\n\tSet objLocator = CreateObject(\"WbemScripting.SWbemLocator\")");
						sHtml.append("\n\tSet objService = objLocator.ConnectServer(strComputer,strNameSpace,strUser,strPass,strLocale,strAuthority,intSecurityFlags)");
						sHtml.append("\n\tobjService.Security_.ImpersonationLevel = intImpersonation");
						sHtml.append("\n\tobjService.Security_.AuthenticationLevel = strAuthentication");
						sHtml.append("\n\tIf intPrivilege = -1 Then");
						sHtml.append("\n\t\tFor p = 1 To 26 Step 1");
						sHtml.append("\n\t\t\tobjService.Security_.Privileges.Add(p)");
						sHtml.append("\n\t\t\tIf Err.Number <> 0 Then");
						sHtml.append("\n\t\t\t\tErr.Clear 'Invalid Parameter");
						sHtml.append("\n\t\t\tEnd If");
						sHtml.append("\n\t\tNext");
						sHtml.append("\n\tElse");
						sHtml.append("\n\t\tobjService.Security_.Privileges.Add(intPrivilege)");
						sHtml.append("\n\tEnd If");
					}
					sHtml.append("\n\tSet objColItems = objService.ExecQuery(\"Select * from " + sClass + "\",\"WQL\",48)");
					sHtml.append("\n\n\tFor Each objItem in objColItems");
					sHtml.append("\n\t\ti = i + 1 : strHtml = \" [INSTANCE-\" & i & \"]\" & Chr(10)");

					for(var oEnum = new Enumerator(oClass.properties_); !oEnum.atEnd(); oEnum.moveNext()){
						var oItem = oEnum.item();
						if(oItem.IsArray){ // Only VBArrays
							/*
							sHtml.append("\n\t\tIf Not IsNull(objItem." + oItem.Name + ") Then ' VBArray object");
							sHtml.append("\n\t\t\tFor a = LBound(objItem." + oItem.Name + ") to UBound(objItem." + oItem.Name + ")");
							sHtml.append("\n\t\t\t\tstrHtml = strHtml & \"  " + oItem.Name + "-\" & a+1 & \": \" & objItem." + oItem.Name + "(a) & Chr(10)");
							sHtml.append("\n\t\t\tNext");
							sHtml.append("\n\t\tElse strHtml = strHtml & \"  " + oItem.Name + ": undefined\" & Chr(10)");
							sHtml.append("\n\t\tEnd If");
							*/
							sHtml.append("\n\t\tstrHtml = strHtml & management_wmi_vbarray(objItem." + oItem.Name + ",\"" + oItem.Name + "\")");
						}
						else if(oItem.CIMType == 101){ // If it is a DateTime object
							/*
							sHtml.append("\n\t\tIf Not IsNull(objItem." + oItem.Name + ") Then ' DateTime object");
							sHtml.append("\n\t\t\tstrDate = Replace(objItem." + oItem.Name + ",\"*\",\"0\")");
							sHtml.append("\n\t\t\tstrDateTime = CDate(Mid(strDate,5,2) & \"/\" & Mid(strDate,7,2) & \"/\" & Left(strDate,4) & \" \" & Mid(strDate,9,2) & \":\" & Mid(strDate,11,2) & \":\" & Mid(strDate,13,2))");
							sHtml.append("\n\t\t\tstrHtml = strHtml & \"  " + oItem.Name + ": \" & strDateTime & Chr(10)");
							sHtml.append("\n\t\tElse strHtml = strHtml & \"  " + oItem.Name + ": undefined\" & Chr(10)");
							sHtml.append("\n\t\tEnd If");
							*/
							sHtml.append("\n\t\tstrHtml = strHtml & \"  " + oItem.Name + ": \" & management_wmi_date(objItem." + oItem.Name + ") & Chr(10)");
						}
						else sHtml.append("\n\t\tstrHtml = strHtml & \"  " + oItem.Name + ": \" & objItem." + oItem.Name + " & Chr(10)");
					}
					sHtml.append("\n\t\tWScript.Echo strHtml");
					sHtml.append("\n\t\tWScript.Sleep 1");
					sHtml.append("\n\tNext");
					sHtml.append("\n\n\tIf Err.Number <> 0 Then");
					sHtml.append("\n\t\tWScript.Echo Err.Description, \"0x\" & Hex(Err.Number)");
					sHtml.append("\n\tEnd If\nEnd Sub");

					// Install Date
					sHtml.append("\n\nFunction management_wmi_date(strDate)");
					sHtml.append("\n\tstrHtml = \"\"");
					sHtml.append("\n\tIf Not IsNull(strDate) Then ' DateTime object");
					sHtml.append("\n\t\tstrDate = Replace(strDate,\"*\",\"0\")");
					sHtml.append("\n\t\tstrDateTime = Left(strDate,4) &  \"-\" & Mid(strDate,5,2) & \"-\" & Mid(strDate,7,2)  & \" \" & Mid(strDate,9,2) & \":\" & Mid(strDate,11,2) & \":\" & Mid(strDate,13,2) & \".\" & Mid(strDate,16,3)");
					sHtml.append("\n\t\tstrHtml = strHtml & strDateTime");
					sHtml.append("\n\tEnd If");
					sHtml.append("\n\tmanagement_wmi_date = strHtml");
					sHtml.append("\nEnd Function");

					// VBArray Function
					sHtml.append("\n\nFunction management_wmi_vbarray(objArray,strArray)");
					sHtml.append("\n\tstrHtml = \"\"");
					sHtml.append("\n\tIf IsArray(objArray) Then ' VBArray object");
					sHtml.append("\n\t\tFor a = LBound(objArray) to UBound(objArray)");
					sHtml.append("\n\t\t\tstrHtml = strHtml & \"  \" & strArray & \"-\" & a+1 & \": \" & objArray(a) & Chr(10)");
					sHtml.append("\n\t\tNext");
					sHtml.append("\n\tElse");
					sHtml.append("\n\t\tstrHtml = strHtml & \"  \" & strArray & \": \" & Chr(10)");
					sHtml.append("\n\tEnd If");
					sHtml.append("\n\tmanagement_wmi_vbarray = strHtml");
					sHtml.append("\nEnd Function");

					//sHtml.append(wmi_vbs_date + "\n\n" + wmi_vbs_vbarray + "\n\n")

					sHtml.append("\n\nWScript.Echo \"D O N E!\"\n ");
				}
				else if(sLanguage.toLowerCase() == "perlscript"){ // PerlScript (Requires ActivePerl for Windows)
					sHtml.append("# Made by: " + __HMBA.mod.purpose + "\n# Generated by: " + document.title);
					sHtml.append("\n# " + __HLang.disclaimer);
					sHtml.append("\n# This Script requires Perl for Windows or ActivePerl at http://www.activeperl.com/Products/Language_Distributions/");
					var h = "\n\nprint \"[ WMI Class: \\\\" + sComputer + "\\" + sNameSpace + ":" + sClass + " ]\";";
					sHtml.append(h.replace(/\\/g,"\\\\"))
					sHtml.append("\n\nuse Win32::OLE qw(in);");
					sHtml.append("\nuse Win32::OLE::Enum;");

					if(bLocal){
						sHtml.append("\n\nsub wmi_" + sClass.toLowerCase() + "($$$$){");
						sHtmlInit = "wmi_" + sClass.toLowerCase() + "(\"" + sComputer + "\",\"" + sNameSpace.replace(/\\/g,"\\\\") + "\",\"" + sImpersonation + "\",\"" + sAuthentication + "\");";
					}
					else{
						sHtml.append("\n\nsub wmi_" + sClass.toLowerCase() + "($$$$$$$$$$){");
						sHtmlInit = "wmi_" + sClass.toLowerCase() + "(\"" + sComputer + "\",\"" + sNameSpace.replace(/\\/g,"\\\\") + "\",\"" + sUser + "\",\"" + sPass + "\",\"" + sImpersonation + "\",\"" + sLocale + "\",\"" + sAuthentication + "\",\"" + sAuthority + "\"," + iPrivilege + "," + iSecurityFlags + ");";
					}

					sHtml.append("\n\tmy $oBool = {0 => \"0\",1 => \"1\", \'\' => \"undefined\"};");
					sHtml.append("\n\tmy $i = 1, $sHtml = \"\";");

					if(bLocal){
						sHtml.append("\n\tmy ($sComputer,$sNameSpace,$sImpersonation,$sAuthentication) = @_;");
						sHtml.append("\n\t$sComputer = SetStr($sComputer,\".\");");
						sHtml.append("\n\t$sNameSpace = SetStr($sNameSpace,\"root\\CIMV2\");".replace(/\\/g,"\\\\"));
						sHtml.append("\n\t$sImpersonation = SetStr($sImpersonation,\"Impersonate\");");
						sHtml.append("\n\t$sAuthentication = SetStr($sAuthentication,\"Connect\");");
						sHtml.append("\n\tmy $oService = Win32::OLE->GetObject(\"winmgmts:{impersonationLevel=\" . $sImpersonation . \",authenticationLevel=\" . $sAuthentication . \",(Security)}!\\\\\\\\\" . $sComputer . \"\\\\\" . $sNameSpace);");
					}
					else {
						sHtml.append("\n\tmy ($sComputer,$sNameSpace,$sUser,$sPass,$sImpersonation,$sLocale,$sAuthentication,$sAuthority,$iPrivilege,$iSecurityFlags) = @_;");
						sHtml.append("\n\t$sComputer = SetStr($sComputer,\".\");");
						sHtml.append("\n\t$sNameSpace = SetStr($sNameSpace,\"root\\CIMV2\");".replace(/\\/g,"\\\\"));
						sHtml.append("\n\t$sUser = SetStr($sUser,undef);");
						sHtml.append("\n\t$sPass = SetStr($sPass,undef);");
						sHtml.append("\n\t$sLocale = SetStr($sLocale,\"MS_409\");");
						sHtml.append("\n\t$sImpersonation = SetStr($sImpersonation,3);");
						sHtml.append("\n\t$sAuthority = SetStr($sAuthority,undef);");
						sHtml.append("\n\t$sAuthentication = SetStr($sAuthentication,4);");
						sHtml.append("\n\t$iSecurityFlags = SetStr($iSecurityFlags,128);");
						sHtml.append("\n\t$iPrivilege = SetStr($iPrivilege,7);");

						sHtml.append("\n\n\tmy $oLocator = Win32::OLE->new(\"WbemScripting.SWbemLocator\");")
						+ " || warn \"Cannot access WMI on local machine: \", Win32::OLE->LastError;";
						sHtml.append("\n\tmy $oService = $oLocator->ConnectServer($sComputer,$sNameSpace,$sUser,$sPass,$sLocale,$sAuthority,$iSecurityFlags,undef)");
						+ " || warn (\"Cannot access WMI on remote machine ($sComputer): \", Win32::OLE->LastError);";
						sHtml.append("\n\t# $oService->Security_->ImpersonationLevel = $sImpersonation; # Error: Can't modify non-lvalue subroutine call");
						sHtml.append("\n\t# $oService->Security_->AuthenticationLevel = $sAuthentication; # Error: Can't modify non-lvalue subroutine call");
						sHtml.append("\n\tif($iPrivilege eq -1){");
						sHtml.append("\n\t\tfor($i = 1; $i <= 26; $i++){");
						sHtml.append("\n\t\t\t$oService->Security_->Privileges->Add($i);");
						sHtml.append("\n\t\t}");
						sHtml.append("\n\t}");
						sHtml.append("\n\telse{");
						sHtml.append("\n\t\t$oService->Security_->Privileges->Add($iPrivilege);");
						sHtml.append("\n\t}");
					}
					sHtml.append("\n\tGetErr();");
					sHtml.append("\n\tmy $oColItems = $oService->ExecQuery(\"Select * from " + sClass + "\",\"WQL\",48);");
					sHtml.append("\n\n\tforeach $oItem (in $oColItems){");
					sHtml.append("\n\t\t$sHtml = \"\\n\\n [INSTANCE-\" . $i++ . \"]\";");
					sHtml.append("\n\t\tif(defined($oItem)){");

					for(var oEnum = new Enumerator(oClass.properties_); !oEnum.atEnd(); oEnum.moveNext()){
						var oItem = oEnum.item();
						if(oItem.IsArray){ // Only VBArrays
							/*
							sHtml.append("\n\t\t\t$s = 1;");
							sHtml.append("\n\t\t\tforeach $a (in @{$oItem->" + oItem.Name + "}){ # Array Object");
							sHtml.append("\n\t\t\t\t$sHtml .= \"\\n  " + oItem.Name + "-\" . $s++ . \": \" . $a;");
							sHtml.append("\n\t\t\t}");
							*/
							sHtml.append("\n\t\t\t$sHtml .= management_wmi_vbarray($oItem->" + oItem.Name + "," + oItem.Name + ");")
						}
						else if(oItem.CIMType == 11){ // If it is a boolean
							sHtml.append("\n\t\t\t$sHtml .= \"\\n  " + oItem.Name + ": \" . $oBool->{$oItem->" + oItem.Name + "};");
						}
						else if(oItem.CIMType == 101){ // If it is a DateTime object
							/*
							sHtml.append("\n\t\t\tif($oItem->" + oItem.Name + " ne ''){ # DateTime Object");
							sHtml.append("\n\t\t\t\t$d = $oItem->" + oItem.Name + ";");
							sHtml.append("\n\t\t\t\t$sDateTime = substr($d,0,4) . \"-\" . substr($d,4,2) . \"-\" . substr($d,6,2) . \" \" . substr($d,8,2) . \":\" . substr($d,10,2) . \":\" . substr($d,12,2) . \".\" . substr($d,15,3);");
							sHtml.append("\n\t\t\t\t$sHtml .= \"\\n  " + oItem.Name + ": \" . $sDateTime;");
							sHtml.append("\n\t\t\t} else {");
							sHtml.append("\n\t\t\t\t$sHtml .= \"\\n  " + oItem.Name + ": undefined\";\n\t\t}");
							*/
							sHtml.append("\n\t\t\t$sHtml .= \"\\n  " + oItem.Name + ": \" . management_wmi_date($oItem->" + oItem.Name + ");");
						}
						else sHtml.append("\n\t\t\t$sHtml .= \"\\n  " + oItem.Name + ": \" . $oItem->" + oItem.Name + ";");

					}
					sHtml.append("\n\t\t}\n\t\tprint $sHtml;\n\t}");
					sHtml.append("\n}");

					// VBDate
					sHtml.append("\n\nsub management_wmi_date($){");
					sHtml.append("\n\tmy $sDateTime = \"\";");
					sHtml.append("\n\tmy($sDate) = @_;");
					sHtml.append("\n\tif(defined($sDate)){ # DateTime Object");
					sHtml.append("\n\t\t$sDate =~ s/\\*/0/g;");
					sHtml.append("\n\t\t$sDateTime = substr($sDate,0,4) . \"-\" . substr($sDate,4,2) . \"-\" . substr($sDate,6,2) . \" \" . substr($sDate,8,2) . \":\" . substr($sDate,10,2) . \":\" . substr($sDate,12,2) . \".\" . substr($sDate,15,3);");
					sHtml.append("\n\t}");
					sHtml.append("\n\treturn $sDateTime;");
					sHtml.append("\n\n}");

					// VBArray
					sHtml.append("\n\nsub management_wmi_vbarray($$){");
					sHtml.append("\n\tmy $sHtml = \"\";");
					sHtml.append("\n\tmy ($oVBarray,$sVBname) = @_;");
					sHtml.append("\n\tif(defined($oVBarray)){");
					sHtml.append("\n\t\t$s = 1;");
					sHtml.append("\n\t\tforeach $a (in @{$oVBarray}){ # Array Object");
					sHtml.append("\n\t\t\t$sHtml .= \"\\n  \" .$sVBname . \"-\" . $s++ . \": \" . $a;");
					sHtml.append("\n\t\t}");
					sHtml.append("\n\t}");
					sHtml.append("\n\telse{");
					sHtml.append("\n\t\t$sHtml .= \"\\n  \" . $sVBname . \": \";");
					sHtml.append("\n\t}");
					sHtml.append("\n\treturn $sHtml;");
					sHtml.append("\n}");

					// GetErr
					sHtml.append("\n\nsub GetErr{");
					sHtml.append("\n\tunless(Win32::OLE->LastError() == 0){");
					sHtml.append("\n\t\tdie(\"Error: \" . Win32::OLE->LastError());");
					sHtml.append("\n\t}");
					sHtml.append("\n}");

					// SetStr
					sHtml.append("\n\nsub SetStr{");
					sHtml.append("\n\tmy $sString = shift;");
					sHtml.append("\n\tmy $sDefault = shift;");
					sHtml.append("\n\tif (not defined($sString)){\n\t\t$sString = $sDefault;\n\t}");
					sHtml.append("\n\treturn $sString;");
					sHtml.append("\n}");

					sHtml.append("\n\n" + sHtmlInit + "\nprint \"\\n\\nD O N E!\";\n");
				}
				else if(sLanguage.toLowerCase() == "wsh"){ // WSH/WSF
					sHtml.append("<?XML version=\"1.0\" standalone=\"yes\" ?>");
					sHtml.append("\n<package>\n\t<job id=\"ILoveJS\">\n\t<?job debug=\"true\"?>");
					sHtml.append("\n\t\t<script language=\"JScript\">\n\t\t<![CDATA[\n");

					var sJSHtml = management_wmi_language(sClass,"jscript",sComputer,sNameSpace,sUser,sPass,sImpersonation,sLocale,sAuthentication,sAuthority,iPrivilege,iSecurityFlags);
					sHtml.append(sJSHtml)

					sHtml.append("\n\t\t]]>\n\t\t</script>");
					sHtml.append("\n\t</job>\n</package>");
				}
				else if(sLanguage.toLowerCase() == "asp_js"){ // ASP
					sHtml.append("<%@ LANGUAGE=\"JSCRIPT\"%>");
					sHtml.append("\n\n<html>\n<!--");
					sHtml.append("\n Made by: " + __HMBA.mod.purpose + "\n Generated by: " + document.title)
					sHtml.append("\n-->\n<%\n/*\n This sample illustrates the use of the WMI Scripting API within an ASP, using JScript.\n It displays information in a table for each disk on the local host. To run this sample:");
					sHtml.append("\n  1. Place it in a directory accessible to your web server\n  2. If running on NT4, ensure that the registry value:\n  HKEY_LOCAL_MACHINE\\Software\\Microsoft\\WBEM\\Scripting\\Enable for ASP is set to 1");
					sHtml.append("\n\n For Windows 2000, ensure that Anonymous Access is disabled and Windows\n Integrated Authentication is enabled for this file before running this ASP\n (this can be done by configuring the file properties using the IIS configuration snap-in).");
					sHtml.append("\n*/\n%>\n\n<head>\n<title></title>\n<style>BODY,TD{font-family:Arial;font-size:11px;padding:2px;vertical-align:top;}</style>\n</head>\n<body scroll=\"auto\">");
					sHtml.append("\n<br><br>\n<table width=\"500\" align=\"center\" cellpadding=\"0\" cellspacing=\"0\" border=\"0\" style=\"background-color:#EEEEEE;border:1px solid #545454;\">");
					sHtml.append("\n<caption align=\"center\" style=\"background-color:black;color:white;\"><b>WMI Class: \\\\" + sComputer + "\\" + sNameSpace + ":" + sClass + "</b></caption>");
					sHtml.append("\n\n<%");
					sHtml.append("\ntry{");
					sHtml.append("\n\tvar sComputer = \"" + (!bLocal ? sComputer : '.') + "\";");
					sHtml.append("\n\tvar sNameSpace = \"" + sNameSpace.replace(/\\/g,"\\\\") + "\";");

					if(bLocal){
						sHtml.append("\n\tvar sImpersonation = \"" + sImpersonation + "\";");
						sHtml.append("\n\tvar sAuthentication = \"" + (sAuthentication ? sAuthentication : "Connect") + "\";");
						sHtml.append("\n\tvar oService = GetObject(\"winmgmts:{impersonationLevel=\" + sImpersonation + \",authenticationLevel=\" + sAuthentication + \",(Security)}!\\\\\\\\\" + sComputer + \"\\\\\" + sNameSpace);");
					}
					else {
						sHtml.append("\n\tvar sUser = \"" + (sUser ? sUser : '') + "\";");
						sHtml.append("\n\tvar sPass = \"" + (sPass ? sPass : '') + "\";");
						sHtml.append("\n\tvar sLocale = \"" + (sLocale ? sLocale : 'MS_409') + "\";");
						sHtml.append("\n\tvar sImpersonation = " + (sImpersonation ? sImpersonation : 3) + ";");
						sHtml.append("\n\tvar sAuthority = \"" + (sAuthority ? sAuthority : "ntlmdomain:" + oWno.UserDomain) + "\";");
						sHtml.append("\n\tvar sAuthentication = \"" + (sAuthentication ? sAuthentication : 4) + "\";");
						sHtml.append("\n\tvar iPrivilege = \"" + (iPrivilege ? iPrivilege : 7) + "\";");
						sHtml.append("\n\tvar iSecurityFlags = \"" + (iSecurityFlags ? iSecurityFlags : 128) + "\";");
						sHtml.append("\n\n\tvar oLocator = new ActiveXObject(\"WbemScripting.SWbemLocator\");");
						sHtml.append("\n\tvar oService = oLocator.ConnectServer(sComputer,sNameSpace,sUser,sPass,sLocale,sAuthority,iSecurityFlags,null);");
						sHtml.append("\n\toService.Security_.ImpersonationLevel = sImpersonation;");
						sHtml.append("\n\toService.Security_.AuthenticationLevel = sAuthentication;");
						sHtml.append("\n\toService.Security_.Privileges.Add(iPrivilege);");
					}
					sHtml.append("\n\tvar oColItems = oService.ExecQuery(\"Select * from " + sClass + "\",\"WQL\",48);");
					sHtml.append("\n\n\tfor(var oItem, i = 1, oEnum = new Enumerator(oColItems); !oEnum.atEnd(); oEnum.moveNext(), i++){");
					sHtml.append("\n\t\toItem = oEnum.item();\n%>");
					sHtml.append("\n\n<tr><td colspan=\"2\">&nbsp;</td></tr>");
					sHtml.append("\n\n<tr><td colspan=\"2\" style=\"background-color:#545454;color:white\"><b>INSTANCE-<%=i%></b></td></tr>");

					for(var oEnum = new Enumerator(oClass.properties_); !oEnum.atEnd(); oEnum.moveNext()){
						var oItem = oEnum.item();
						if(oItem.IsArray){ // Only VBArrays
							sHtml.append("\n\n<%");
							sHtml.append("\n\t\tif(oItem." + oItem.Name + " != null){");
							sHtml.append("\n\t\t\tvar a" + oItem.Name + " = new VBArray(oItem." + oItem.Name + ").toArray();");
							sHtml.append("\n\t\t\tfor(var j = 0; j < a" + oItem.Name + ".length; j++){");
							sHtml.append("\n%>");
							sHtml.append("\n\n<tr>\n\t<td>" + oItem.Name + "-<%=(j+1)%></td>\n\t<td>&nbsp;<%=a" + oItem.Name + "[j]%></td>\n</tr>");
							sHtml.append("\n\n<%");
							sHtml.append("\n\t\t\t}");
							sHtml.append("\n\t\t}");
							sHtml.append("\n\t\telse {\n%>\n\n<tr>\n\t<td>" + oItem.Name + "</td>\n\t<td>&nbsp;undefined</td>\n</tr>");
							sHtml.append("\n\n<%\n\t\t}\n%>\n");
						}
						else if(oItem.CIMType == 101){ // If it is a DateTime object
							sHtml.append("\n\n<%");
							sHtml.append("\n\t\tif(oItem." + oItem.Name + " != null){");
							sHtml.append("\n\t\t\tvar d = oItem." + oItem.Name + ";");
							sHtml.append("\n\t\t\tvar sDateTime = d.substring(0,4) + \"-\" + d.substring(4,6) + \"-\" + d.substring(6,8) + \" \" + d.substring(8,10) + \":\" + d.substring(10,12) + \":\" + d.substring(12,14);");
							sHtml.append("\n%>\n");
							sHtml.append("\n<tr>\n\t<td>" + oItem.Name + "</td>\n\t<td>&nbsp;<%=sDateTime%></td>\n</tr>");
							sHtml.append("\n\n<%");
							sHtml.append("\n\t\t}");
							sHtml.append("\n\t\telse {\n%>\n<tr>\n\t<td>" + oItem.Name + "</td>\n\t<td>&nbsp;undefined</td>\n</tr>");
							sHtml.append("\n\n<%\n\t\t}\n%>\n");
						}
						else sHtml.append("\n<tr>\n\t<td>" + oItem.Name + "</td>\n\t<td>&nbsp;<%=oItem." + oItem.Name + "%></td>\n</tr>");
					}
					sHtml.append("\n\n<%\n\t}\n}");
					sHtml.append("\ncatch(e){");
					sHtml.append("\n%>\n\nError:<br>Number: <%=(e.number & 0xFFFF)%><br>Description: <%=e.description%>");
					sHtml.append("\n\n<%\n}\n%>");
					sHtml.append("\n\n</table>\n\n</body>\n</html>\n");
				}
				else if(sLanguage.toLowerCase() == "kixstart"){ // KiXstart
					sHtml.append("; Made by: " + __HMBA.mod.purpose + "\n; Generated by: " + document.title);
					sHtml.append("\n; " + __HLang.disclaimer);
					sHtml.append("\n; http://msdn.microsoft.com/library/en-us/wmisdk/wmi/win32_classes.asp")
					sHtml.append("\n\nBREAK ON");
					sHtml.append("\n\n;Redirect Outputs: kix32.exe " + sClass + ".kix $sReDirect='redirect.txt'")
					sHtml.append("\nIf IsDeclared($sReDirect) $r = ReDirectOutput($sReDirect,0) EndIf")
					sHtml.append("\n\n? \"[ WMI Class: \\\\" + sComputer + "\\" + sNameSpace + ":" + sClass + " ]\" + Chr(10)");

					if(bLocal){
						sHtml.append("\n\nwmi_" + sClass.toLowerCase() + "(\"" + sComputer + "\",\"" + sNameSpace + "\",\"" + sImpersonation + "\",\"" + sAuthentication + "\")");
						sHtml.append("\n\nFunction wmi_" + sClass.toLowerCase() + "($sComputer,$sNameSpace,$sImpersonation,$sAuthentication)");
					}
					else{
						sHtml.append("\n\nwmi_" + sClass.toLowerCase() + "(\"" + sComputer + "\",\"" + sNameSpace + "\",\"" + sUser + "\",\"" + sPass + "\",\"" + sImpersonation + "\",\"" + sLocale + "\",\"" + sAuthentication + "\",\"" + sAuthority + "\"," + iPrivilege + "," + iSecurityFlags + ")");
						sHtml.append("\n\nFunction wmi_" + sClass.toLowerCase() + "($sComputer,$sNameSpace,$sUser,$sPass,$sImpersonation,$sLocale,$sAuthentication,$sAuthority,$iPrivilege,$iSecurityFlags)");
					}

					sHtml.append("\n\tDim $oService, $oLocator")
					sHtml.append("\n\t$sHtml = \"\"\n\t$i = 0");

					sHtml.append("\n\n\t$sComputer = SetStr($sComputer,'.')");
					sHtml.append("\n\t$sNameSpace = SetStr($sNameSpace,'root\\CIMV2')");
					if(bLocal){
						sHtml.append("\n\t$sImpersonation = SetStr($sImpersonation,'Impersonate')");
						sHtml.append("\n\t$sAuthentication = SetStr($sAuthentication,'Connect')");
						sHtml.append("\n\t$oService = GetObject(\"winmgmts:{impersonationLevel=\" + $sImpersonation + \",authenticationLevel=\" + $sAuthentication + \",(Security)}!\\\\\" + $sComputer + \"\\\" +  $sNameSpace)");
					}
					else {
						sHtml.append("\n\t$sUser = SetStr($sUser,'')");
						sHtml.append("\n\t$sPass = SetStr($sPass,'')");
						sHtml.append("\n\t$sLocale = SetStr($sLocale,'MS_409')");
						sHtml.append("\n\t$iImpersonation = SetInt($iImpersonation,3)");
						sHtml.append("\n\t$sAuthentication = SetInt($sAuthentication,4)");
						sHtml.append("\n\t$sAuthority = SetStr($sAuthority,'')");
						sHtml.append("\n\t$iSecurityFlags = SetInt($iSecurityFlags,128)");
						sHtml.append("\n\t$iPrivilege = SetInt($iPrivilege,7)");

						sHtml.append("\n\n\t$oLocator = CreateObject(\"WbemScripting.SWbemLocator\")");
						sHtml.append("\n\t$oService = $oLocator.ConnectServer($sComputer,$sNameSpace,$sUser,$sPass,$sLocale,$sAuthority,$iSecurityFlags)");
						sHtml.append("\n\t$oService.Security_.ImpersonationLevel = $iImpersonation");
						sHtml.append("\n\t$oService.Security_.AuthenticationLevel = $sAuthentication");
						sHtml.append("\n\tIf $iPrivilege == -1");
						sHtml.append("\n\t\tFor $p = 1 To 26 Step 1");
						sHtml.append("\n\t\t\t$oService.Security_.Privileges.Add($p)");
						sHtml.append("\n\t\t\tGetErr(False,True)");
						sHtml.append("\n\t\tNext");
						sHtml.append("\n\tElse");
						sHtml.append("\n\t\t$oService.Security_.Privileges.Add($iPrivilege)");
						sHtml.append("\n\tEndIf");
					}
					sHtml.append("\n\n\t$oColItems = $oService.ExecQuery(\"Select * from " + sClass + "\",\"WQL\",48)");
					sHtml.append("\n\tGetErr()");

					sHtml.append("\n\n\tFor Each $oItem in $oColItems");
					sHtml.append("\n\t\t$i = $i + 1\n\t\t$sHtml = \" [INSTANCE-\" + $i + \"]\" + Chr(10)");

					for(var oEnum = new Enumerator(oClass.properties_); !oEnum.atEnd(); oEnum.moveNext()){
						var oItem = oEnum.item();
						if(oItem.IsArray){ // Only VBArrays
							sHtml.append("\n\t\t$sHtml = $sHtml + management_wmi_vbarray($oItem." + oItem.Name + ",\"" + oItem.Name + "\")");
						}
						else if(oItem.CIMType == 11){ // If it is a boolean
							sHtml.append("\n\t\t$sHtml = $sHtml + \"  " + oItem.Name + ": \" + GetBoolean($oItem." + oItem.Name + ") + Chr(10)");
						}
						else if(oItem.CIMType == 101){ // If it is a DateTime $oect
							sHtml.append("\n\t\t$sHtml = $sHtml + \"  " + oItem.Name + ": \" + management_wmi_date($oItem." + oItem.Name + ") + Chr(10)");
						}
						else sHtml.append("\n\t\t$sHtml = $sHtml + \"  " + oItem.Name + ": \" + $oItem." + oItem.Name + " + Chr(10)");
					}
					sHtml.append("\n\t\t? $sHtml");
					sHtml.append("\n\tNext");
					sHtml.append("\n\tGetErr()");
					sHtml.append("\nEndFunction");

					// Install Date
					sHtml.append("\n\nFunction management_wmi_date($sDate)");
					sHtml.append("\n\t$sHtml = \"\"");
					sHtml.append("\n\tIf VarType($sDate) == 12 Or VarType($sDate) == 8; String or DateTime Object");
					sHtml.append("\n\t\t;$sDate = Replace($sDate,\"*\",\"0\");Not implemented");
					sHtml.append("\n\t\t$sDateTime = Left($sDate,4) + \"-\" + SubStr($sDate,5,2) + \"-\" + SubStr($sDate,7,2) + \" \" + SubStr($sDate,9,2) + \":\" + SubStr($sDate,11,2) + \":\" + SubStr($sDate,13,2) + \".\" + SubStr($sDate,16,3)");
					sHtml.append("\n\t\t$sHtml = $sHtml + $sDateTime");
					sHtml.append("\n\tEndIf");
					sHtml.append("\n\t$management_wmi_date = $sHtml");
					sHtml.append("\nEndFunction");

					// VBArray Function
					sHtml.append("\n\nFunction management_wmi_vbarray($oArray,$sArray)");
					sHtml.append("\n\t$sHtml = \"\"");
					sHtml.append("\n\tIf VarType($oArray) >= 8192; VBArray Object");
					sHtml.append("\n\t\tFor $a = 0 to UBound($oArray)");
					sHtml.append("\n\t\t\t$sHtml = $sHtml + \"  \" + $sArray + \"-\" + ($a+1) + \": \" + $oArray[$a] + Chr(10)");
					sHtml.append("\n\t\tNext");
					sHtml.append("\n\tElse");
					sHtml.append("\n\t\t$sHtml = $sHtml + \"  \" + $sArray + \": \" + Chr(10)");
					sHtml.append("\n\tEndIf");
					sHtml.append("\n\t$management_wmi_vbarray = $sHtml");
					sHtml.append("\nEndFunction");

					// SetStr Function
					sHtml.append("\n\nFunction SetStr($sString,$sDefault)");
					sHtml.append("\n\t$sString = Trim($sString)");
					sHtml.append("\n\tIf $sString = ''");
					sHtml.append("\n\t\t$sString = $sDefault");
					sHtml.append("\n\tEndIf");
					sHtml.append("\n\t$SetStr = $sString");
					sHtml.append("\nEndFunction");

					// SetInt
					sHtml.append("\n\nFunction SetInt($iInteger,$iDefault)");
					sHtml.append("\n\tIf VarType($iInteger) <> 3");
					sHtml.append("\n\t\t$iInteger = $iDefault");
					sHtml.append("\n\tEndIf");
					sHtml.append("\n\t$SetInt = $iInteger");
					sHtml.append("\nEndFunction");

					// GetBoolean
					sHtml.append("\n\nFunction GetBoolean($iBoolean)");
					sHtml.append("\n\t$sBoolean = \"False\"");
					sHtml.append("\n\tIf VarType($iBoolean) == 3 And $iBoolean == -1");
					sHtml.append("\n\t\t$sBoolean = \"True\"");
					sHtml.append("\n\tEndIf");
					sHtml.append("\n\t$GetBoolean = $sBoolean");
					sHtml.append("\nEndFunction");

					// GetErr
					sHtml.append("\n\nFunction GetErr(optional $bExit, optional $bClear)");
					sHtml.append("\n\tIf @ERROR <> 0");
					sHtml.append("\n\t\t? \" Number: \" + @ERROR + Chr(10) + \" Description: \" + @SERROR  + Chr(10)");
					sHtml.append("\n\t\tIf $bExit\n\t\t\tExit If(@ERROR < 0, VAL(\"&\" + Right(DecToHex(@ERROR),4)),@ERROR)\n\t\tEndIf");
					sHtml.append("\n\tEndIf");
					sHtml.append("\n\tIf $bClear @ERROR = 0 EndIf");
					sHtml.append("\nEndFunction");

					sHtml.append("\n\n? \"D O N E!\"\n");
				}
				else if(sLanguage.toLowerCase() == "rexx"){ // Object rexx http://www.mindspring.com/~dave_martin/RexxFAQ.html
					sHtml.append("/* Made by: " + __HMBA.mod.purpose + "\n* Generated by: " + document.title);
					sHtml.append("\n* " + __HLang.disclaimer);
					sHtml.append("\n* This Script requires IBM Object Rexx for Windows http://www.rexx.com\n*/");
					sHtml.append("\n\nsay '[ WMI Class: \\\\" + sComputer + "\\" + sNameSpace + ":" + sClass + " ]'");
					sHtml.append("\n\n");

					if(bLocal){
						sHtml.append("wmi_" + sClass.toLowerCase() + ":\nProcedure Expose sComputer sNameSpace sImpersonation sAuthentication");
						sHtmlInit = "wmi_" + sClass.toLowerCase() + "(\"" + sComputer + "\",\"" + sNameSpace.replace(/\\/g,"\\\\") + "\",\"" + sImpersonation + "\",\"" + sAuthentication + "\")";
					}
					else{
						sHtml.append("wmi_" + sClass.toLowerCase() + ":\nProcedure Expose sComputer sNameSpace sUser sPass sImpersonation sAuthentication sLocale sAuthentication sAuthority iPrivilege iSecurityFlags");
						sHtmlInit = "wmi_" + sClass.toLowerCase() + "(\"" + sComputer + "\",\"" + sNameSpace.replace(/\\/g,"\\\\") + "\",\"" + sUser + "\",\"" + sPass + "\",\"" + sImpersonation + "\",\"" + sLocale + "\",\"" + sAuthentication + "\",\"" + sAuthority + "\"," + iPrivilege + "," + iSecurityFlags + ")";
					}

					sHtml.append("\n\ti = 1\n\tsHtml = \"\"");

					if(bLocal){
						sHtml.append("\n\tsComputer = \".\"");
						sHtml.append("\n\tsNameSpace = \"root\\CIMV2\"");
						sHtml.append("\n\toService = .OLEObject.GetObject(\"winmgmts:\\\\\"||sComputer||\"\\\"||sNameSpace)");
					}
					else {
						sHtml.append("\n\tsComputer = \".\"");
						sHtml.append("\n\tsNameSpace = \"root\\CIMV2\"");
						sHtml.append("\n\toService = .OLEObject~ConnectServer(sComputer,sNameSpace)");

					}
					sHtml.append("\n\toColItems = oService~ExecQuery(\"Select * from " + sClass + "\",\"WQL\",48)");
					sHtml.append("\n\n\tdo oItem over oColItems");
					sHtml.append("\n\t\tsHtml = \"\\n\\n [INSTANCE-\" i++ \"]\"");

					for(var oEnum = new Enumerator(oClass.properties_); !oEnum.atEnd(); oEnum.moveNext()){
						var oItem = oEnum.item();
						if(oItem.IsArray){ // Only VBArrays
							sHtml.append("\n\t\tsHtml = sHtml management_wmi_vbarray(oItem~" + oItem.Name + ",\"" + oItem.Name + "\")");
						}
						/*
						else if(oItem.CIMType == 11){ // If it is a boolean
						sHtml.append("\n\t\t\tsHtml = sHtml || \"\\n  " + oItem.Name + ": \" + oBool.{oItem~" + oItem.Name + "}");
						}
						*/
						else if(oItem.CIMType == 101){ // If it is a DateTime object
							sHtml.append("\n\t\tsHtml = sHtml \"\\n  " + oItem.Name + ": \" management_wmi_date(oItem~" + oItem.Name + ")");
						}
						else sHtml.append("\n\t\tsHtml = sHtml \"\\n  " + oItem.Name + ": \" oItem~" + oItem.Name + "");

					}
					sHtml.append("\n\t\tsay sHtml\n\treturn");

					// VBDate
					sHtml.append("\n\nmanagement_wmi_date:\nProcedure oDtm");
					sHtml.append("\n\tsDateTime = \"\"");
					sHtml.append("\n\tif(oDtm[4] == 0):");
					sHtml.append("\n\t\tsDateTime = oDtm[5] + '-'");
					sHtml.append("\n\telse:");
					sHtml.append("\n\t\tsDateTime = oDtm[4] + oDtm[5] + '-'");
					sHtml.append("\n\tif(oDtm[6] == 0):");
					sHtml.append("\n\t\tsDateTime = sDateTime + oDtm[7] + '-'");
					sHtml.append("\n\telse:");
					sHtml.append("\n\t\tsDateTime = sDateTime + oDtm[6] + oDtm[7] + '-'");
					sHtml.append("\n\t\tsDateTime = oDtm[0] + oDtm[1] + oDtm[2] + oDtm[3] + sDateTime + \" \" + oDtm[8] + oDtm[9] + \":\" + oDtm[10] + oDtm[11] + \":\" + oDtm[12] + oDtm[13]");
					sHtml.append("\n\treturn sDateTime");

					// VBArray
					sHtml.append("\n\nmanagement_wmi_vbarray:\nProcedure oVBarray sVBname");
					sHtml.append("\n\tsItem = \"\"");
					sHtml.append("\n\ts = 1");
					sHtml.append("\n\tif .NIL <> oVBarray then");
					sHtml.append("\n\t\tdo x over oVBarray");
					sHtml.append("\n\t\t\tsItem = sItem \"\\n  \" sVBname \"-\" s++ \": \" x");
					sHtml.append("\n\tend");
					sHtml.append("\n\treturn sItem");

					sHtml.append("\n\n" + sHtmlInit + "\n\nsay \"\\n\\nD O N E!\"\n");
				}
				else if(sLanguage.toLowerCase() == "python"){ // Python
					sHtml.append("# Made by: " + __HMBA.mod.purpose + "\n# Generated by: " + document.title);
					sHtml.append("\n# " + __HLang.disclaimer);
					sHtml.append("\n# This Script requires Python for Windows http://www.python.org/doc");
					sHtml.append("\n\nprint \"[ WMI Class: \\\\" + sComputer + "\\" + sNameSpace + ":" + sClass + " ]\"");
					sHtml.append("\n\nimport win32com.client");

					if(bLocal){
						sHtml.append("\n\ndef wmi_" + sClass.toLowerCase() + "(sComputer,sNameSpace,sImpersonation,sAuthentication):");
						sHtmlInit = "wmi_" + sClass.toLowerCase() + "(\"" + sComputer + "\",\"" + sNameSpace.replace(/\\/g,"\\\\") + "\",\"" + sImpersonation + "\",\"" + sAuthentication + "\")";
					}
					else{
						sHtml.append("\n\ndef wmi_" + sClass.toLowerCase() + "(sComputer,sNameSpace,sUser,sPass,sImpersonation,sAuthentication,sLocale,sAuthentication,sAuthority,iPrivilege,iSecurityFlags):");
						sHtmlInit = "wmi_" + sClass.toLowerCase() + "(\"" + sComputer + "\",\"" + sNameSpace.replace(/\\/g,"\\\\") + "\",\"" + sUser + "\",\"" + sPass + "\",\"" + sImpersonation + "\",\"" + sLocale + "\",\"" + sAuthentication + "\",\"" + sAuthority + "\"," + iPrivilege + "," + iSecurityFlags + ")";
					}

					sHtml.append("\n\ti = 1, sHtml = \"\"");

					if(bLocal){
						sHtml.append("\n\tsComputer = \".\"");
						sHtml.append("\n\tsNameSpace = sNameSpace #SetStr(sNameSpace)");
						sHtml.append("\n\toService = win32com.client.Dispatch(\"WbemScripting.SWbemLocator\")");
						sHtml.append("\n\toService = oService.ConnectServer(sComputer,sNameSpace)");
					}
					else {
						sHtml.append("\n\tsComputer = \".\"");
						sHtml.append("\n\tsNameSpace = sNameSpace #SetStr(sNameSpace)");
						sHtml.append("\n\toService = win32com.client.Dispatch(\"WbemScripting.SWbemLocator\")");
						sHtml.append("\n\toService = oService.ConnectServer(sComputer,sNameSpace)");

						/*
						sHtml.append("\n\t$sComputer = SetStr($sComputer,\".\");");
						sHtml.append("\n\t$sNameSpace = SetStr($sNameSpace,\"root\\CIMV2\");".replace(/\\/g,"\\\\"))
						sHtml.append("\n\t$sUser = SetStr($sUser,undef);");
						sHtml.append("\n\t$sPass = SetStr($sPass,undef);");
						sHtml.append("\n\t$sLocale = SetStr($sLocale,\"MS_409\");");
						sHtml.append("\n\t$sImpersonation = SetStr($sImpersonation,3);");
						sHtml.append("\n\t$sAuthority = SetStr($sAuthority,undef);");
						sHtml.append("\n\t$sAuthentication = SetStr($sAuthentication,4);");
						sHtml.append("\n\t$iSecurityFlags = SetStr($iSecurityFlags,128);");
						sHtml.append("\n\t$iPrivilege = SetStr($iPrivilege,7);");

						sHtml.append("\n\n\tmy $oLocator = Win32::OLE->new(\"WbemScripting.SWbemLocator\");");
							+ " || warn \"Cannot access WMI on local machine: \", Win32::OLE->LastError;";
						sHtml = sHtml + "\n\tmy $oService = $oLocator->ConnectServer($sComputer,$sNameSpace,$sUser,$sPass,$sLocale,$sAuthority,$iSecurityFlags,undef)"
							+ " || warn (\"Cannot access WMI on remote machine ($sComputer): \", Win32::OLE->LastError);";
						sHtml.append("\n\t# $oService->Security_->ImpersonationLevel = $sImpersonation; # Error: Can't modify non-lvalue subroutine call");
						sHtml.append("\n\t# $oService->Security_->AuthenticationLevel = $sAuthentication; # Error: Can't modify non-lvalue subroutine call");
						sHtml.append("\n\tif($iPrivilege eq -1){");
						sHtml.append("\n\t\tfor($i = 1; $i <= 26; $i++){");
						sHtml.append("\n\t\t\t$oService->Security_->Privileges->Add($i);");
						sHtml.append("\n\t\t}");
						sHtml.append("\n\t}");
						sHtml.append("\n\telse{");
						sHtml.append("\n\t\t$oService.Security_->Privileges->Add($iPrivilege);");
						sHtml.append("\n\t}");
						*/
					}
					sHtml.append("\n\toColItems = oService.ExecQuery(\"Select * from " + sClass + "\",\"WQL\",48)");
					sHtml.append("\n\n\tfor oItem in oColItems:");
					sHtml.append("\n\t\tsHtml = \"\\n\\n [INSTANCE-\" + i++ + \"]\"");

					for(var oEnum = new Enumerator(oClass.properties_); !oEnum.atEnd(); oEnum.moveNext()){
						var oItem = oEnum.item();
						if(oItem.IsArray){ // Only VBArrays
							sHtml.append("\n\t\tsHtml = sHtml + management_wmi_vbarray(oItem." + oItem.Name + ",\"" + oItem.Name + "\")");
						}
						/*
						else if(oItem.CIMType == 11){ // If it is a boolean
						sHtml.append("\n\t\t\tsHtml = sHtml + \"\\n  " + oItem.Name + ": \" + oBool.{oItem->" + oItem.Name + "}");
						}
						*/
						else if(oItem.CIMType == 101){ // If it is a DateTime object
							sHtml.append("\n\t\tsHtml = sHtml + \"\\n  " + oItem.Name + ": \" + management_wmi_date(oItem." + oItem.Name + ")");
						}
						else sHtml.append("\n\t\tsHtml = sHtml + \"\\n  " + oItem.Name + ": \" + oItem." + oItem.Name + "");

					}
					sHtml.append("\n\t\tprint sHtml");

					// VBDate
					sHtml.append("\n\ndef management_wmi_date(oDtm):");
					sHtml.append("\n\tsDateTime = \"\"");
					sHtml.append("\n\tif(oDtm[4] == 0):");
					sHtml.append("\n\t\tsDateTime = oDtm[5] + '-'");
					sHtml.append("\n\telse:");
					sHtml.append("\n\t\tsDateTime = oDtm[4] + oDtm[5] + '-'");
					sHtml.append("\n\tif(oDtm[6] == 0):");
					sHtml.append("\n\t\tsDateTime = sDateTime + oDtm[7] + '-'");
					sHtml.append("\n\telse:");
					sHtml.append("\n\t\tsDateTime = sDateTime + oDtm[6] + oDtm[7] + '-'");
					sHtml.append("\n\t\tsDateTime = oDtm[0] + oDtm[1] + oDtm[2] + oDtm[3] + sDateTime + \" \" + oDtm[8] + oDtm[9] + \":\" + oDtm[10] + oDtm[11] + \":\" + oDtm[12] + oDtm[13]");
					sHtml.append("\n\treturn sDateTime");

					// VBArray
					sHtml.append("\n\ndef management_wmi_vbarray(oVBarray,sVBname):");
					sHtml.append("\n\tsItem = \"\"");
					sHtml.append("\n\ttry:");
					sHtml.append("\n\t\ts = 1");
					sHtml.append("\n\t\tfor a in oVBarray:");
					sHtml.append("\n\t\t\tsItem = sItem + \"\\n  \" + sVBname + \"-\" + s++ + \": \" + a");
					sHtml.append("\n\texcept:");
					sHtml.append("\n\t\tsItem = sItem + \"\\n  \" + sVBname + \": \"");
					sHtml.append("\n\treturn sItem");

					sHtml.append("\n\n" + sHtmlInit + "\n\nprint \"\\n\\nD O N E!\"\n");
				}
				return sHtml.toString();
			}
			catch(e){
				__HLog.error(e,this);
				return false;
			}
			finally{
				sHtml.clear()
			}
		}