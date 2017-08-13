// Copyright© 2003-2008. nOsliw Solutions. All rights reserved.
//**Start Encode**

__H.register(__H.UI.Window.HTA.MBA._Module,"DBDocumentLLD","Module Database Document LLD",function DBDocumentLLD(){
	
	var bIsLoaded = false
	var a_qry_title = []
	var s_qry_title = "select Title from LLD_Documents"
	var a_qry_chapter = []
	var s_qry_chapter = "select * from LLD_Chapters order by Number"
	var a_qry_content = []
	var s_qry_content = "select * from LLD_Contents"
	
	
	var db_access_file = __H.s_pth_root + "\\" + __HMBA.mdl_cur.FileDB
	
	this.init = function init(){
		try{
			if(bIsLoaded) return;
			__H.include("HUI@Sys@ADO.js")
			
			if(__HADO.connectAccess(db_access_file,__HMBA.mdl_cur.FileDBUser,__HMBA.mdl_cur.FileDBPass)){
				// Title
				a_qry_title = __HADO.selectArray(s_qry_title)
				__HSelect.setSelect(oFormDBDocumentLLD.lld_title)
				__HSelect.addArray(a_qry_title["Title"],1)
				// Chapters
				this.create()
			}
			else {
				__HLog.logPopup("# Unable to load Access database file: " + db_access_file)
			}
			bIsLoaded = true
		}
		catch(ee){
			__HLog.error(ee,this);
			return false
		}
	}
	
	this.create = function create(){
		try{
			oDivDBDocumentLLD2.innerHTML = ""
			a_qry_chapter = __HADO.selectArray(s_qry_chapter)
			for(var i = 0, iLen = a_qry_chapter["Name"].length; i < iLen; i++){
				var s = __H.byClone("span")
				s.className = "cSpanHeader1_1"
				s.innerHTML = a_qry_chapter["Number"][i] + " " + a_qry_chapter["Name"][i] + "<br>"
				oDivDBDocumentLLD2.appendChild(s)
				
				if(a_qry_chapter["Header"][i]){
					s.innerHTML = "<br>" + s.innerHTML + "<br>"
					s.className = "cSpanHeader1"
					continue
				}
				var d = __H.byClone("div")		
				d.contentEditable = true
				d.className = "cDivChapter"
				d.innerHTML = "<br>"				
				
				oDivDBDocumentLLD2.appendChild(d)				
			}
		}
		catch(ee){
			__HLog.error(ee,this);
			return false
		}
	}
	
	this.update = function update(){
		try{
			if(!bIsLoaded) return;
			a_qry_content = __HADO.selectArray(s_qry_content)
			
		}
		catch(ee){
			__HLog.error(ee,this);
			return false
		}
	}
	
	this.update2 = function update2(){
		try{
			
		}
		catch(ee){
			__HLog.error(ee,this);
			return false
		}
	}
	
	this.update3 = function update3(){
		try{
			
		}
		catch(ee){
			__HLog.error(ee,this);
			return false
		}
	}
	
})

var __HLLD = new __H.UI.Window.HTA.MBA._Module.DBDocumentLLD()

function db_document_init(){
	try{
		//var sFile = "L:\\Development\\HTA\\MBA\\1.0\\module\\nosliw\\database\\data\\db\\hui-database-document.mdb"
		
		return true
	}
	catch(e){
		__HLog.error(e,this);
		return false
	}
}

function database_doc_(){
	try{
		__HLog.logPopup("You clicked the blue button")
		
		return true
	}
	catch(e){
		__HLog.error(e,this);
		return false
	}
}

function database_doc(){
	try{
		__HLog.popup("You run Example-1 for the first time")
		
		return true
	}
	catch(e){
		__HLog.error(e,this);
		return false
	}
}