// Copyright© 2003-2008. nOsliw Solutions. All rights reserved.
//**Start Encode**


__H.register(__H.UI.Window.HTA.MBA._Module,"Notepad","Class for Notepad",function Notepad(){
	
	this.variable1_protected = null
	var variable_private     = null
	variable_global          = null
	var bIsInitialized       = false
	
	this.inititialize = function inititialize(){
		
		
		bIsInitialized = true
		return true
	}
	
	this.isExample = function isExample(){
		// TODO:
		return true
	}
	
	this.open = function open(){
		if(!this.isExample()) return;
		try{
			var f = __HShell.browseFile()
			alert(f)
			return true
		}
		catch(ee){
			__HLog.error(ee,this)
			return false
		}
	}
	
	this.save = function save(){
		if(!this.isExample()) return;
		try{
			
			return true
		}
		catch(ee){
			__HLog.error(ee,this)
			return false
		}
	}
	
	this.saveas = function saveas(){
		if(!this.isExample()) return;
		try{
			
			return true
		}
		catch(ee){
			__HLog.error(ee,this)
			return false
		}
	}
	
	this.close = function close(){
		try{
			
			return true
		}
		catch(ee){
			__HLog.error(ee,this)
			return false
		}
	}
})

var __HNotepad = new __H.UI.Window.HTA.MBA._Module.Notepad()