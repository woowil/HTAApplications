// nOsliw HUI - HTML/HTA Application Framework Library (http://hui.codeplex.com/)
// Copyright (c) 2003-2013, nOsliw Solutions,  All Rights Reserved.
// License: GNU Library General Public License (LGPL) (http://hui.codeplex.com/license)
//**Start Encode**

__H.register(__H.UI.Window.HTA.MBA,"_Style","Class for HTA style",function _Style(){
	var menu_bordercolor		= "#336600"//"#316AC5"
	var menu_backgroundcolor	= "#CCCC99"//"#C1D2EE"
	var menubar_bordercolor		= "#8A867A"//"#ECE9D8"
	var menubar_backgroundcolor = "#EFEDDE"
	var pic_path = "addons/skins/officexp/"
	var previous_tag = null
	
	var isIE7 = function(){
		var o = navigator.userAgent
		return o.indexOf("MSIE 7") || o.indexOf("MSIE 6") || o.indexOf("MSIE 5.5")	// "MSIE 5" has bugs	
	}
	
	var isColor = function(sColor){
		if(typeof(sColor) != "string") return false
		return sColor.isSearch(/#{0,1}[a-z0-9]{6}$/ig)
	}
	
	this.setSkinPath = function setSkinPath(sPath){
		if(oFso.FolderExists(sPath)){
			pic_path = sPath
		}
	}
	
	this.getSkinPath = function getSkinPath(){
		return pic_path
	}
	
	this.setColor = function setColor(sMenuBorder,sMenuBack,sMenubarBorder,sMenubarBack){
		try{//alert(isColor(sMenubarBack))
			menu_bordercolor		= isColor(sMenuBorder) ? sMenuBorder : menu_bordercolor
			menu_backgroundcolor	= isColor(sMenuBack) ? sMenuBack : menu_backgroundcolor
			menubar_bordercolor		= isColor(sMenubarBorder) ? sMenubarBorder : menubar_bordercolor
			menubar_backgroundcolor	= isColor(sMenubarBack) ? sMenubarBack : menubar_backgroundcolor
			return true
		}
		catch(ee){
			__HLog.error(ee,this)
			return false
		}
	}
		
	this.setMenuColor = function setMenuColor(oMenu,sMenuBorder,sMenuBack,sMenubarBorder,sMenubarBack){
		try{
			//this.test()
			if(this.setColor(sMenuBorder,sMenuBack,sMenubarBorder,sMenubarBack)){
				if(typeof(oMenu) == "object"){
					//this.menuBarItem(oMenu) // causes the menu to flipp
					//this.test()
					return true
				}
			}
			return false
		}
		catch(ee){
			__HLog.error(ee,this)
			return false
		}
	}
	
	this.test = function(){
		var background	= "url(" + pic_path + "divider.png) " + menubar_backgroundcolor + " no-repeat left"
		var filter		= "progid:DXImageTransform.Microsoft.Gradient(GradientType=1,StartColorStr='" + menubar_backgroundcolor + "',EndColorStr='#CCCC99')";
		
		alert(background)
	}
	
	////////////////// MENUBAR
	///////////////////////////////////////////////////////////////
	
	this.menuBarItem = function menuBarItem(oElement,bChange){
		try{
			with(oElement.style){
				if(!bChange){
					cursor		= "default"
					borderStyle = "outset"
					borderWidth = "1px 0 1px 1px"
					padding		= "1px 1px 6px 1px"
					position	= "absolute"
					visibility	= "visible"
					top			= 0
					left		= 0
					width		= "auto"
					height		= isIE7() ? "auto" : "1%"
					zIndex		= 100
				}				
				background	= "url(" + pic_path + "divider.png) " + menubar_backgroundcolor + " no-repeat left"
				filter		= "progid:DXImageTransform.Microsoft.Gradient(GradientType=1,StartColorStr='" + menubar_backgroundcolor + "',EndColorStr='#CCCC99')";
			}
		}
		catch(ee){
			__HLog.error(ee,this)
		}
	}
	
	this.menuBarItemDrag = function menuBarItemDrag(oElement){
		try{
			with(oElement.style){
				cursor		= "move"
				display		= "inline"
				font		= "11px Arial normal"
				padding		= "3px 1px 3px 1px"
				position	= "relative"
				visibility	= "hidden"
				width		= "9px"
			}
		}
		catch(ee){
			__HLog.error(ee,this)
		}
	}
	
	this.menuBarItemOut = function menuBarItemOut(oElement){
		try{
			with(oElement.style){
				backgroundColor	= menubar_backgroundcolor
				borderWidth = "0"
				color		= "#000000"
				display		= "inline"
				font		= "11px Arial normal"
				padding		= "2px 5px 2px 5px"
				position	= "relative"
			}
		}
		catch(ee){
			__HLog.error(ee,this)
		}
	}
	
	this.menuBarItemOver = function menuBarItemOver(oElement){
		try{
			if(previous_tag) this.menuBarItemOut(previous_tag)
			with(oElement.style){
				backgroundColor	= menu_backgroundcolor
				borderColor = menu_bordercolor
				borderStyle = "solid"
				borderWidth = "1px 0 1px 0"
			}
			previous_tag = oElement
		}
		catch(ee){
			__HLog.error(ee,this)
		}
	}
	
	this.menuBarItemClick = function menuBarItemClick(oElement){
		try{
			with(oElement.style){
				backgroundColor	= menubar_backgroundcolor
				borderColor = menubar_bordercolor
			}
		}
		catch(ee){
			__HLog.error(ee,this)
		}
	}
	
	////////////////// MENU
	///////////////////////////////////////////////////////////////	
	
	this.menu = function menu(oElement){
		try{
			with(oElement.style){
				background	  = "#FFFFFF url(" + pic_path + "menu_left.png) repeat-y"
				border		  = "1px solid #8A867A"
				color		  = "#000000"
				cursor		  = "default"
				paddingBottom = "1px"
				paddingTop	  = "1px"
				position	  = "absolute"
				visibility	  = "hidden"
				zIndex		  = "100"
			}
		}
		catch(ee){
			__HLog.error(ee,this)
		}
	}
	
	this.menuLayer = function menuLayer(oElement){
		try{
			with(oElement.style){
				background	= "transparent"
				display		= "none"
				filter		= "progid:DXImageTransform.Microsoft.dropshadow(OffX=5,OffY=5,Color='gray',Positive='true')"
				position	= "absolute"
				visibility	= "hidden"
				zIndex		= "90"
			}
		}
		catch(ee){
			__HLog.error(ee,this)
		}
	}
	
	this.menuItemOut = function menuItemOut(oElement){
		try{
			with(oElement.style){
				background	= "transparent"
				border		= "none"
				color		= "#000000"
				font		= "11px Arial normal"
				margin		= "0"
				padding		= "3px 15px 3px 30px"
				position	= "relative"
			}
		}
		catch(ee){
			__HLog.error(ee,this)
		}
	}
	
	this.menuItemOver = function menuItemOver(oElement){
		try{
			with(oElement.style){
				backgroundColor = menu_backgroundcolor
				borderColor		= menu_bordercolor
				borderStyle		= "solid"
				borderWidth		= "1px 0 1px 0"
				padding			= "2px 15px 2px 30px"
			}
		}
		catch(ee){
			__HLog.error(ee,this)
		}
	}
	
	this.menuItemClick = function menuItemClick(oElement){
		try{
			with(oElement.style){
				
			}
		}
		catch(ee){
			__HLog.error(ee,this)
		}
	}
	
	this.menuItemSep = function menuItemSep(oElement){
		try{
			with(oElement.style){
				paddingLeft = "28px"
			}
		}
		catch(ee){
			__HLog.error(ee,this)
		}
	}
	
	////////////////// ICONS
	///////////////////////////////////////////////////////////////
		
	this.menuIconOut = function menuIconOut(oElement){
		try{
			oElement.className = "mba-toolbar-icon-1" 
			/*
			with(oElement.style){
				background	= "url(icon1.png) no-repeat"
				height		= "16px"
				left		= "4px"
				position	= "absolute"
				width		= "16px"
			}*/
		}
		catch(ee){
			__HLog.error(ee,this)
		}
	}
	
	this.menuIconOver = function menuIconOver(oElement){
		try{
			oElement.className = "mba-toolbar-icon-2" 
			/*
			with(oElement.style){
				background	= "url(icon2.png) no-repeat"
			}*/
		}
		catch(ee){
			__HLog.error(ee,this)
		}
	}
	
	this.menuIconClick = function menuIconClick(oElement){
		try{
			oElement.className = "mba-toolbar-icon-3" 
			/*
			with(oElement.style){
				background	= "url(icon3.png) no-repeat"
			}*/
		}
		catch(ee){
			__HLog.error(ee,this)
		}
	}
	
	this.menuArrowOut = function menuArrowOut(oElement){
		try{
			with(oElement.style){
				background	= "url(" + pic_path + "arrow_right.gif) no-repeat"
				height		= "7px"
				position	= "absolute"
				right		= "8px"
				width		= "4px"
			}
		}
		catch(ee){
			__HLog.error(ee,this)
		}
	}
	
	this.menuArrowOver = function menuArrowOver(){};
})

var __MStyle = new __H.UI.Window.HTA.MBA._Style()
