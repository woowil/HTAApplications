// Copyright© 2003-2008. nOsliw Solutions. All rights reserved.
//**Start Encode**

__H.include("HUI@IO@File_JSCompression.js")

__H.register(__H.UI.Window.HTA.MBA._Module,"Utilities","Utilities",function Utilities(){
	var oForm

	this.password = function password(){
		try{
			__HLog.logInfo("# Getting random passwords..")
			oForm = oFormUtilityPasswd
			var linearCongruential = function(seed){ // X(n+1) = (aX(n) + c) mod m, n >= 0. http://www.merriampark.com/lcm.htm
				var aPrimes1 = [1000,1024,1009,7], aPrimes2 = [5,11,101,71,19]
				var B = aPrimes2[(0).random(aPrimes2.length)]
				var M = aPrimes1[(0).random(aPrimes1.length)]
				var A = (new Date()).getDate()
				return (A*seed + B) % M
			};
			var o = {}
			var s = new __H.Common.StringBuffer()
			var p = new __H.Common.StringBuffer()
			o.number = "0123456789"
			o.alpha_upper = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
			o.alpha_lower = "abcdefghijklmnopqrstuvwxyz"
			o.none_ascii = "!@@#$%^&*-_+=[]\;',./`~(){}|:<>?"
			o.unicode = "åæøÅÆØäöÄÖ"
			
			s.append(oForm.passwd_number.checked ? o.number : "")
			s.append(oForm.passwd_alphaupper.checked ? o.alpha_upper : "")
			s.append(oForm.passwd_alphalower.checked ? o.alpha_lower : "")
			s.append(oForm.passwd_nonascii.checked ? o.none_ascii : "")
			s.append(oForm.passwd_unicode.checked ? o.unicode : "")
			s.append(oForm.passwd_string.value)
			
			if(oForm.passwd_shuffle.checked) s.shuffle()
			if(oForm.passwd_encode.checked) s.encode()
			var v1 = parseInt(oForm.passwd_quantity.value)
			var v2 = parseInt(oForm.passwd_length.value)
			for(var i = 0; i < v1; i++){
				p.append((i+1) + ". ")
				for(var j = 0; j < v2; j++){
					//if() linearCongruential(j+i)
					p.append(s.random())
				}
				p.append("<br>")
			}
			//var oElement = __H.byIds("oDivUtilityPasswdResult")		
			oDivUtilityPasswdResult.innerHTML = p
			__HHTA.onScroll(oDivUtilityPasswdResult)
			s.empty(), p.empty()
			return false // No need to refresh
		}
		catch(e){
			__HLog.error(e,this);
			return false;
		}
	}


	this.compress = function compress(){
		try{
			//__HLog.logInfo("# Getting random passwords..")
			oForm = oFormUtilityCompress
			var bFile = oForm.compress_file[0].checked
			var bCompress = oForm.compress_type[0].checked		
			var sSrcFolder = oFso.GetParentFolderName(oForm.src_folder.value)
			var sDstFolder = oFso.GetParentFolderName(oForm.dst_folder.value)
			var sSrcFile = oForm.src_file.value
			if(__H.isStringEmpty(sDstFolder)){
				__HLog.popup("Destination folder is needed!")
				return false;
			}
			else if(sSrcFolder == sDstFolder){
				__HLog.popup("Source and destination must be different!")
				return false
			}
			
			__HCompression.initialize({
				b_delcomments	: oForm.compress_delcomments.checked,
				b_dellines		: oForm.compress_dellines.checked,
				b_addsemicolon	: oForm.compress_addsemicolon.checked,
				b_mergelines	: oForm.compress_mergelines.checked,
				b_delwhitespace : oForm.compress_delwhitespace.checked,
				b_recursive		: oForm.compress_recursive.checked,
				b_encrypt		: oForm.compress_encrypt.checked	
			})
			if(bFile){ // File
				if(bCompress) __HCompression.compressFile(sSrcFile,sDstFolder)
				else __HCompression.decompressFile(sSrcFile,sDstFolder)
			}
			else {
				if(bCompress) __HCompression.compressFolder(sSrcFolder,sDstFolder)
				else __HLog.popup("Not implemented!");
			}
			__HShell.open(sDstFolder)
			
			return true
		}
		catch(e){
			__HLog.error(e,this);
			return false;
		}
	}
})

var __HUtilities = new __H.UI.Window.HTA.MBA._Module.Utilities()


function utility_mail(sOpt){
	var oOutlook, oMapi, oMail;
	try{ // http://www.outlookcode.com/ http://www.outlookcode.com/d/code/htmlimg.htm
		if(!__HMBA.mod.sendmail){
			__HLog.info(__HMBA.txt.deactivated,__HMBA.txt.deactivated_css);
			oFormAction.action[0].disabled = true
			return;
		}
		oOutlook = new ActiveXObject("Outlook.Application");
		if(sOpt == "getcontacts"){	
 			alert("M@il will now try to get a list of current contact in your Outlook.")
 			oMapi = oOutlook.GetNamespace("MAPI")
  			var objContactFolder = oMapi.GetDefaultFolder(10);
  			var sHTML = '<select name="mail_to"><option>' 
  			//Populate the pulldown
  			for(var i = 1; i <  objContactFolder.Items.Count; i++){
    			sHTML = sHTML.concat("<option value='" +
    			objContactFolder.Items(i).Email1Address + "'>" +
    			objContactFolder.Items(i).FullName)
  			}
  			sHTML = sHTML.concat("</select>")
  			// Insert the pulldown into the SPAN area in below HTML
  			oSpanUtilMailContacts.insertAdjacentHTML("beforeEnd", sHTML);
		}
		else if(sOpt == "sendmail"){
			oMail = oOutlook.CreateItem(olMailItem)
    		oMail.To = oFormUtilityMail.mail_to.value
    		oMail.Subject = oFormUtilityMail.mail_subject.value
    		oMail.Body = oFormUtilityMail.mail_message.value
    		if(!__H.isUndefined(oMail.To,oMail.Subject,oMail.Body)) oMail.Send();
    		else __HLog.popup("Missing subject or message");
		}
		return true
	}
	catch(e){
		if((e.number & 0xFFFF) == 429){ // Automation server can't create object
			__HLog.logPopup("# Unable to load Outlook MAPI on current system. M@il will be deactivated.");
			oFormAction.action[0].disabled = true
			__HMBA.mod.sendmail = false
		}
		__HLog.error(e,this);
		return false;
	}
	finally{
		__HUtil.kill(oOutlook,oMapi,oMail);
	}
}

function utility_cleartbody(sTBody){
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
