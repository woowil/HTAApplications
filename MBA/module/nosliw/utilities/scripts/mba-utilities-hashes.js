// Copyright© 2003-2008. nOsliw Solutions. All rights reserved.
//**Start Encode**

var oHashX, oHmacX, oZmeyX;

function utility_hashes_(){
	try{
		
	}
	catch(e){
		__HLog.error(e,this);
		return false;
	}
}

function utility_hashes_init(){
	try{
		__HMBA.fls.hashes = __HMBA.mdl_cur.FileBin + "\\hashes.dll";
		__HMBA.hashes = [];
			__HMBA.hashes["ZmeYsoft.Hashes.Hash"] = new ObjectZmey("ZmeYsoft.Hashes.Hash","","","Universal HASH object");
			__HMBA.hashes["ZmeYsoft.Hashes.CRC32"] = new ObjectZmey("ZmeYsoft.Hashes.CRC32","CRC32",32,"Cycle Redundancy Check");
			__HMBA.hashes["ZmeYsoft.Hashes.ADLER32"] = new ObjectZmey("ZmeYsoft.Hashes.ADLER32","ADLER32",32,"Adler-32 checksum");
			__HMBA.hashes["ZmeYsoft.Hashes.MD4"] = new ObjectZmey("ZmeYsoft.Hashes.MD4","MD4",128,"RSA Data Security. MD4 Message-digest algorithm");
			__HMBA.hashes["ZmeYsoft.Hashes.MD5"] = new ObjectZmey("ZmeYsoft.Hashes.MD5","MD5",128,"RSA Data Security. MD5 Message-digest algorithm");
			__HMBA.hashes["ZmeYsoft.Hashes.SHA1"] = new ObjectZmey("ZmeYsoft.Hashes.SHA1","SHA1",160,"Secure Hash Algorithm (rev. 1)");
			__HMBA.hashes["ZmeYsoft.Hashes.RIPEMD128"] = new ObjectZmey("ZmeYsoft.Hashes.RIPEMD128","RIPEMD128",128,"RIPEMD-128");
			__HMBA.hashes["ZmeYsoft.Hashes.RIPEMD160"] = new ObjectZmey("ZmeYsoft.Hashes.RIPEMD160","RIPEMD160",160,"RIPEMD-160");
			__HMBA.hashes["ZmeYsoft.Hashes.RIPEMD256"] = new ObjectZmey("ZmeYsoft.Hashes.RIPEMD256","RIPEMD256",256,"RIPEMD-256 (same level as RIPEMD-128)");
			__HMBA.hashes["ZmeYsoft.Hashes.RIPEMD320"] = new ObjectZmey("ZmeYsoft.Hashes.RIPEMD320","RIPEMD320",320,"RIPEMD-320 (same level as RIPEMD-160");
		__HMBA.hmac = [];
			__HMBA.hmac["ZmeYsoft.HMACs.HMAC"] = new ObjectZmey("ZmeYsoft.HMACs.HMAC","","","Universal HMAC object");
			__HMBA.hmac["ZmeYsoft.HMACs.MD5"] = new ObjectZmey("ZmeYsoft.HMACs.MD5","MD5",128,"RSA Data Security. MD5 Message-digest algorithm");
			__HMBA.hmac["ZmeYsoft.HMACs.SHA1"] = new ObjectZmey("ZmeYsoft.HMACs.SHA1","SHA1",160,"Secure Hash Algorithm (rev. 1)");
			__HMBA.hmac["ZmeYsoft.HMACs.RIPEMD128"] = new ObjectZmey("ZmeYsoft.HMACs.RIPEMD128","HMAC-RIPEMD128",128,"Uses RIPEMD-128");
			__HMBA.hmac["ZmeYsoft.HMACs.RIPEMD160"] = new ObjectZmey("ZmeYsoft.HMACs.RIPEMD160","HMAC-RIPEMD160",160,"Uses RIPEMD-160");		
		__HMBA.zmey = {};
			__HMBA.zmey.hash = false;
			__HMBA.zmey.hash_idx = false; // Either Hashes or Hmac
			__HMBA.zmey.question = __HMBA.zmey.file = __HMBA.zmey.folder = __HMBA.zmey.key = "";
			__HMBA.zmey.idx = false;
		return true;
	}
	catch(e){
		__HLog.error(e,this);
		return false;
	}
}

function utility_hashes_load(){
	try{
		if(!utility_hashes_init()) return false;
		var oForm = oFormUtilityHash
		
		__HLog.log("# Registering ActiveX Library ZmeYsoft.Hashes.Hash");
		if(!(oHashX = __HSys.regActiveX("ZmeYsoft.Hashes.Hash",__HMBA.fls.hashes))){
			__HLog.log("## Error loading/registering ActiveX Library ZmeYsoft.Hashes.Hash\n\nTry restarting the application manually again!");
			return false;
		}
		__HLog.log("# Registering ActiveX Library ZmeYsoft.HMACs.HMAC");
		if(!(oHmacX = __HSys.regActiveX("ZmeYsoft.HMACs.HMAC",__HMBA.fls.hashes))){
			__HLog.log("## Error loading/registering ActiveX Library ZmeYsoft.HMACs.HMAC\n\nTry restarting the application manually again!");
			return false;
		}
				
		// HASHES
		var j = 0, b;
		__HSelect.setSelect(oForm.hashes)
		for(var sHash in __HMBA.hashes){
			if(!__HMBA.hashes.hasOwnProperty(sHash)) continue
			if(j++ == 0) continue
			var sBits = (((b = __HMBA.hashes[sHash].bits) < 100) ? "0" + b : b) + " / ";			
			__HSelect.addIndex(null,__HMBA.hashes[sHash].progid,sBits + __HMBA.hashes[sHash].algorithm)
		}
		
		// HMAC		
		j = 0;
		__HSelect.setSelect(oForm.hmac)
		for(var sHmac in __HMBA.hmac){
			if(!__HMBA.hmac.hasOwnProperty(sHmac)) continue
			if(j++ == 0) continue
			var sBits = (((b = __HMBA.hmac[sHmac].bits) < 100) ? "0" + b : b) + " / ";			
			__HSelect.addIndex(null,__HMBA.hmac[sHmac].progid,sBits + __HMBA.hmac[sHmac].algorithm)
		}
		
		oForm.hashes.onchange = oForm.hmac.onchange = new Function("utility_hashes_formact(1,this)");
		oForm.operation1[0].onclick = oForm.operation1[1].onclick = new Function("utility_hashes_formact(2,this)");
		
		oForm.operation2[0].onclick = oForm.operation2[1].onclick = oForm.operation2[2].onclick = new Function("utility_hashes_formact(2,this)");
		oForm.question.onchange = oForm.file.onchange = oForm.folder.onchange = new Function("utility_hashes_formact(2,this)");
		
		__HSelect.setSelect()
		return true;
	}
	catch(e){
		__HLog.error(e,this);
		return false;
	}
}

function utility_hashes_main(sOpt){
	try{
		var oHash, sHash = new __H.Common.StringBuffer();
		if(sOpt == "hash"){
			__HLog.logInfo("# Getting XmeY Soft Hash")
			if(!__H.isStringEmpty(__HMBA.zmey.hash) && __HMBA.zmey.idx){				
				if(oHash = utility_hashes_hash(__HMBA.zmey.idx,__HMBA.zmey.question,__HMBA.zmey.file,__HMBA.zmey.folder,true)){
					var oBody = __H.byIds("oBdyUtilityHash"), oRow, oCell, n
					oRow = oBody.insertRow(oBody.rows.length), oRow.style.verticalAlign = 'top', n = (oBody.rows.length).toNumberZero();
					oCell = oRow.insertCell(), oCell.innerText = n, oCell.width = 22, oCell.align = "center"
					for(var o in oHash){
						if(!oHash.hasOwnProperty(o)) continue
						sHash.append("Hash " + o + ": " + oHash[o] + "\n")
					}
					oCell = oRow.insertCell(), oCell.innerText = sHash.toString()
					oCell = oRow.insertCell(), oCell.innerText = n, oCell.width = 22, oCell.align = "center"
					return true
				}
				else __HLog.logPopup("## Hash failed or illegal Hash String!")
			}
			else __HLog.logPopup("## Either the BITS/ALGORITHM is not selected or illegal Hash String!")
		}
		return false
	}
	catch(e){
		__HLog.error(e,this);		
		return false;
	}
}

function ObjectZmey(sProgid,sAlgorithm,iBits,sDescription){
	this.progid = sProgid;
	this.algorithm = sAlgorithm;
	this.bits = iBits;
	this.description = sDescription;
}

function utility_hashes_formact(iOpt,oElement){
	try{
		if(iOpt == 1){
			__HSelect.setSelect(oElement)
			var sAlgorithm = __HMBA.zmey.zmey = __HSelect.getValue()			
			if(!sAlgorithm) return;
			if(oElement.name == "hashes") {
				oDivHashProgId.innerHTML = sAlgorithm;
				oDivHashDesc.innerHTML = __HMBA.hashes[sAlgorithm].description;
				__HMBA.zmey.hash = __HMBA.hashes[sAlgorithm].algorithm;
				oZmeyX = oHashX;
			}
			else if(oElement.name == "hmac") {
				oDivHmacProgId.innerHTML = sAlgorithm;
				oDivHmacDesc.innerHTML = __HMBA.hmac[sAlgorithm].description;
				__HMBA.zmey.hash = __HMBA.hmac[sAlgorithm].algorithm;
				oZmeyX = oHmacX;
			}
			if(__HMBA.zmey.hash_idx == 2) __HMBA.zmey.key = oElement.form.key.value
		}
		else if(iOpt == 2){ // 
			__HMBA.zmey.otag = oElement;
			if(oElement.value == "hashes"){
				oElement.form.hmac.disabled = oElement.form.key.disabled = true;
				oElement.form.hashes.disabled = false;
				__HMBA.zmey.hash_idx = 1;
			}
			else if(oElement.value == "hmac"){
				oElement.form.hmac.disabled = oElement.form.key.disabled = false;
				oElement.form.hashes.disabled = true;
				__HMBA.zmey.hash_idx = 2;
			}
			else if(oElement.value == "question" || oElement.name == "question"){
				oElement.form.question.disabled = false;
				oElement.form.file.disabled = oElement.form.folder.disabled = true;
				__HMBA.zmey.question = oElement.value;
				__HMBA.zmey.idx = 1;
			}
			else if(oElement.value == "file" || oElement.name == "file"){
				oElement.form.file.disabled = false;
				oElement.form.question.disabled = oElement.form.folder.disabled = true;
				__HMBA.zmey.file = oElement.value;
				__HMBA.zmey.idx = 2;
			}
			else if(oElement.value == "folder" || oElement.name == "folder"){
				oElement.form.folder.disabled = false;
				oElement.form.file.disabled = oElement.form.question.disabled = true;
				oElement.value = __HMBA.zmey.folder = oFso.GetParentFolderName(oElement.value);
				__HMBA.zmey.idx = 3;
			}	 
		}
		else if(iOpt == 3){	
			__HMBA.zmey.otag.value = __HMBA.zmey.folder;
			
		}
		else if(iOpt == 9){ // Scroll Textarea
			if(oElement.scrollHeight > oElement.clientHeight){
				oElement.scrollTop = (oElement.scrollHeight-oElement.clientHeight);
			}
		}
	}
	catch(e){		
		__HLog.error(e,this);
		return;
	}
}

function utility_hashes_hash(iOpt,sSecret,sFile,sFolder,bClone){
	try{
		var oFile, oHash = false;
		if(iOpt == 1 && !__H.isStringEmpty(sSecret)){ // Hash Question
			oHash = {};
			oZmeyX.HashAnsiString(sSecret);
			oHash.Value = oZmeyX.Value;
			oZmeyX.Algorithm = __HMBA.zmey.hash;
			oHash.Algorithm = oZmeyX.Algorithm;
			oHash.Binary = oZmeyX.BinaryValue;
			oHash.Secret = sSecret;
			oHash.File = oHash.Folder = "";
			oHash.ZmeY = __HMBA.zmey.zmey;
			oHash.Key = (__HMBA.zmey.idx == 2) ? __HMBA.zmey.key : "";
			oZmeyX.Reset();
		}
		else if(iOpt == 2 && __HFile.exists(sFile)){ // Hash file
			oHash = {};
			oZmeyX.HashFile(oFso.GetAbsolutePathName(sFile));
			oHash.Value = oZmeyX.Value;
			oHash.Algorithm = oZmeyX.Algorithm;
			oHash.Binary = oZmeyX.BinaryValue;
			oHash.Secret = oHash.Folder = "";
			oHash.File = sFile;
			oHash.ZmeY = __HMBA.zmey.zmey;
			oHash.Key = (__HMBA.zmey.idx == 2) ? __HMBA.zmey.key : "";
			oZmeyX.Reset();			
		}
		else if(iOpt == 3 && !__H.isStringEmpty(sFolder)){ // Hash folder
			if(aFiles = __HFile.listFiles(sFolder)){
				var sValue = "";
				oHash = {};
				for(var i = 0, iLen = aFiles.length; i < iLen; i++){
					//__HLog.log(aFiles[i])
					oZmeyX.HashFile(aFiles[i]);
					if(bClone) sValue = oZmeyX.Clone().Value; // Doesn't work!!
					else { // On each file (concat string)
						sValue = sValue.concat(oZmeyX.Value)
						oZmeyX.Reset();
					}
				}
				oHash.Value = sValue;
				oHash.Algorithm = oZmeyX.Algorithm;
				oHash.Binary = oZmeyX.BinaryValue;
				oHash.Secret = oHash.File = "";
				oHash.Folder = sFolder;
				oHash.ZmeY = __HMBA.zmey.zmey;
				oHash.Key = (__HMBA.zmey.idx == 2) ? __HMBA.zmey.key : "";				
				oZmeyX.Reset();
			}
		}
		return oHash;
	}
	catch(e){
		__HLog.error(e,this);
		return false;
	}
}
