// Copyright© 2003-2008. nOsliw Solutions. All rights reserved.
//**Start Encode**

function desktop_search(){
	try{
		__HLog.popup("Hello World! This is a popup")
		__HLog.log("Hello World! This is a log message")
		
		return true
	}
	catch(e){
		__HLog.error(e,this);
		return false
	}
}

function desktop_search2(){
	try{
		__HLog.logPopup("You clicked the blue button")
		
		return true
	}
	catch(e){
		__HLog.error(e,this);
		return false
	}
}

function desktop_search3(){
	try{
		__HLog.popup("You run Example-1 for the first time")
		
		return true
	}
	catch(e){
		__HLog.error(e,this);
		return false
	}
}