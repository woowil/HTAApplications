// Copyright© 2003-2008. nOsliw Solutions. All rights reserved.
//**Start Encode**

//__H.include("HUI@Unknown@Class@Library.js","HUI@Another@Class@Library.js")

__H.register(__H.UI.Window.HTA.MBA._Module,"ExampleFunction","Class for Unknown..",function ExampleFunction(){
	var o_this = this
	var b_initialized = false
	
	this.variable1_protected = null
	var variable_private     = null
	variable_global          = null
		
	var d_templates
	var o_template
	
	/////////////////////////////////////
	//// DEFAULT
	
	var o_options = {
		say : "Hello World"
	}
	
	var initialize = function initialize(bForce){
		if(b_initialized && !bForce) return;

		__HLog.debug("Initalizing class: __H.My.Parent.Class._Template")
		
		d_templates = new ActiveXObject("Scripting.Dictionary")
		o_template  = {}
		
		b_initialized = true
	}
	
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
	
	/////////////////////////////////////
	//// 
	
		
	this.isExample = function isExample(){
		// TODO:
		return true
	}
	
	this.getMethod = function getMethod(){
		if(!this.isExample()) return;
		try{
			
			return true
		}
		catch(ee){
			__HLog.error(ee,this)
			return false
		}
	}
	
	this.getMethod = function getMethod(){
		if(!this.isExample()) return;
		try{
			
			return true
		}
		catch(ee){
			__HLog.error(ee,this)
			return false
		}
	}

	var private_function = function private_function(){
		try{
			
			return true
		}
		catch(ee){
			__HLog.error(ee,this)
			return false
		}
	}
	
	global_function = function global_function(){
		try{
			
			return true
		}
		catch(ee){
			__HLog.error(ee,this)
			return false
		}
	}
})

// Global Alias variable for your class
// var __HExpFunc = new __H.UI.Window.HTA.MBA._Module.ExampleFunction()

function example_unknown_function(){
	try{
		__HLog.popup("Hello World! This is a popup")
		__HLog.log("Hello World! This is a log message")
		
		return true
	}
	catch(ee){
		__HLog.error(ee,this);
		return false
	}
}

function example_unknown_function2(){
	try{
		__HLog.logPopup("This is a popup and an log message section2")
		
		return true
	}
	catch(ee){
		__HLog.error(ee,this);
		return false
	}
}

function example_unknown_function3(){
	try{
		__HLog.debug("This is a debug log message section3")
		
		return true
	}
	catch(ee){
		__HLog.error(ee,this);
		return false
	}
}