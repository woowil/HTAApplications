// Copyright© 2003-2008. nOsliw Solutions. All rights reserved.
//**Start Encode**

function management_process(sOpt){
	try{
		var oForm = oFormManagementProcess
		var oBody = __H.byIds("oBdyManagementProcess")	
		
		if(sOpt == "refresh"){
			for(var oEnum = new Enumerator(__HMBA.wmi.u_colitems), r = 0; !oEnum.atEnd(); oEnum.moveNext()){
				var oItem = oEnum.item()
				if(r > 0) r++ // skip seperator row
				for(var i = 0; i < __HMBA.wmi.u_properties.length; i++, r++){
					var o = __HMBA.wmi.u_properties[i], oCell, c
					
					if(o.isarray) c = (c=__HWMI.wmiArray(oItem[o.name],o.name,"\n  ")) ? c.stream : ""
					else if(o.isdatetime) c = __HWMI.wmiDate(oItem[o.name])
					else c = oItem[o.name]
					o.isnumber = (typeof(c) == "number") ? true : false					
					
					oCell = oBody.rows(r).cells(3), oCell.innerText = c // Current Process
					oCell = oBody.rows(r).cells(4) // Average
					if(o.isnumber) oCell.u_total += c, oCell.u_denom++, oCell.innerText = Math.ceil(oCell.u_total/oCell.u_denom)
					else oCell.innerText = 0		
					oCell = oBody.rows(r).cells(5) // Maximum
					if(o.isnumber) oCell.innerText = oCell.u_max = Math.ceil(Math.max(c,oCell.u_max))
					else oCell.innerText = 0
					oCell = oBody.rows(r).cells(6) // Minimum
					if(o.isnumber) oCell.innerText = oCell.u_min = Math.ceil(Math.min(c,oCell.u_min))
					else oCell.innerText = 0
				}
			}
			if(oForm.process_interval_stop.checked){
				oForm.process_classes.disabled = true
				__HUtil.sleep(oForm.process_interval.value)				
				__HWMI.setRefresher(__HMBA.wmi.u_class,__HMBA.wmi.u_server)
				__HMBA.wmi.u_process_id = setTimeout("management_process('refresh')",0)
			}
			else {
				clearTimeout(__HMBA.wmi.u_process_id), __HMBA.wmi.u_process_id = null
				oForm.process_classes.disabled = false
			}
		}
		else {
			var oCap = __H.byIds("oCapManagementProcess")
			__HMBA.wmi.u_class  = oForm.process_classes.value
			
			__HLog.logInfo("# Getting Process Monitor on system: " + __HMBA.wmi.u_server)
			
			__HMBA.wmi.u_colitems = __HWMI.getRefresher(__HMBA.wmi.u_class,__HMBA.wmi.u_server);
			if(!__HMBA.wmi.u_colitems){
				__HLog.logPopup("## Unable to make a WMI connection to system: " + __HMBA.wmi.u_server);
				return false;
			}
			
			__HMBA.wmi.u_properties = __HWMI.getProperty(__HMBA.wmi.u_class)
			oCap.innerHTML = "<i>WMI Process Monitor on</i> <b>" + __HMBA.wmi.u_server + "</b> using class <b>" + __HMBA.wmi.u_class + "</b>"
			
			management_cleartbody(oBody)
			
			for(var oEnum = new Enumerator(__HMBA.wmi.u_colitems), iSep = 0; !oEnum.atEnd(); oEnum.moveNext()){
				var oItem = oEnum.item(), oRow
				if(iSep++ > 0){
					oRow = oBody.insertRow(oBody.rows.length)
					oCell = oRow.insertCell(), oCell.innerText = " "
					oCell.className = "mba-head-2", oCell.colSpan = 8
				}
				for(var i = 0; i < __HMBA.wmi.u_properties.length; i++){
					oRow = oBody.insertRow(oBody.rows.length), oCell
					oRow.style.verticalAlign = 'top';
					var o = __HMBA.wmi.u_properties[i], n = (oBody.rows.length).toNumberZero()
					
					oCell = oRow.insertCell(), oCell.innerText = n, oCell.width = 22, oCell.align = "center"
					oCell = oRow.insertCell(), oCell.innerText = o.name, oCell.noWrap = true, oCell.width = "1%", oCell.className = "mba-table-1TD"
					
					if(o.isarray) v = (v=__HWMI.wmiArray(oItem[o.name],o.name,"\n  ")) ? v.stream : ""
					else if(o.isdatetime) v = __HWMI.wmiDate(oItem[o.name])
					else v = oItem[o.name]
					o.isnumber = (typeof(v) == "number") ? true : false
					
					oCell = oRow.insertCell(), oCell.innerText = o.typestr
					oCell = oRow.insertCell(), oCell.innerText = v, oCell.bgColor = "#EEFFEE" // Current
					oCell = oRow.insertCell()
					if(o.isnumber) oCell.innerText = oCell.u_total = v, oCell.u_denom = 1// Bug. First value can't be a number!
					else oCell.innerText = v
					oCell = oRow.insertCell(), oCell.innerText = oCell.u_max = o.isnumber ? v : 0// Maximum
					oCell = oRow.insertCell(), oCell.innerText = oCell.u_min = o.isnumber ? v : 0 // Minimum					
					
					oCell = oRow.insertCell(), oCell.innerText = n, oCell.width = 22, oCell.align = "center"
				}
			}
			setTimeout("oFormAction.action[1].onclick()",4000)
			return true
		}
		return false
	}
	catch(e){
		__HLog.error(e,this);
		return false;
	}
}

function management_cleartbody(sTBody){
	try{
		var oBody = typeof(sTBody) == "object" ? sTBody : __H.byIds(sTBody)
		for(var i = oBody.rows.length; i > 0; i--) oBody.deleteRow(i-1);
		return true
	}
	catch(e){
		__HLog.error(e,this);
		return false
	}
}