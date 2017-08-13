// Copyright© 2003-2006. nOsliw Solutions. All rights reserved.
//**Start Encode**

var ani

__H.register(__H.UI.Window.HTA.MBA._Module,"GIFAnimate","Modules Games GIF Animate",function GIFAnimate(x,y){
	// PROTECTED VARIABLES
	var load = false
	var anime_img = []
	var anime_num = 0
	var anime_len = 0
	anime_ID = null
	var interval = 50
	var delay = 100
	var delay_MAX = 4000		
	
	this.objImg = __H.byIds("oImgAnimate")
	
	this.init = function(){
		if(!load){
			var aImages = []
			for(var i = 0; i < 10; i++){ //__HMBA.pth.pics_url
				aImages[i] = "module/nosliw/games/data/pic/cartoon-fight-" + (i+1) + ".gif"					
			}
			this.load(aImages,140,104)
		}
		this.initStart = true
		this.start()
		this.initStart = false
	}
				
	// PRIVATE FUNCTION
	this.load = function(aImages,w,h){
		if(typeof(aImages) != "object" || !aImages.length) return false
		anime_img.length = 0
		for(var i = 0, iLen = aImages.length; i < iLen; i++){
			anime_img[i] = (w && h) ? new Image(w,h) : new Image();
			anime_img[i].src = aImages[i] // Remember to use front slash: "bin-pics/cartoon-fight" + i + ".gif";
		}
		load = (i > 0)
		anime_len = aImages.length
		return load
	}

	// PRIVATE FUNCTION
	this.start = function start(){
		try{
			if(this.initStart || (typeof(obj) == "object" && typeof(this.oload) != "object")){
				this.oload = oFormAction.action[0]
				this.oslower = oFormAction.action[1]
				this.ostop = oFormAction.action[2]
				this.ostart = oFormAction.action[3]
				this.ofaster = oFormAction.action[4]		
			}
			if(this.ostart){ // Only for oFormAction
				this.ostart.disabled = this.oload.disabled = true
				this.ostop.disabled = this.oslower.disabled = this.ofaster.disabled = false
			}
			if(!anime_ID && load){
				anime_ID = window.setTimeout(__HMBA.gamesAnime(),interval);
			}
		}
		catch(ee){
			this.error(ee,this)
		}
	}

	// PRIVATE FUNCTION
	this.stop = function(){
		if(this.ostop){ // Only for oFormAction
			this.ostart.disabled = false
			this.ostop.disabled = this.oslower.disabled = this.ofaster.disabled = this.oload.disabled = true
		}
		if(anime_ID){
			clearTimeout(anime_ID);
		}
		anime_ID = null
	}

	// PUBLIC FUNCTION
	__HMBA.gamesAnime = function(){
		if(anime_ID && load){
			if(anime_num >= anime_len) anime_num = 0
			ani.objImg.src = anime_img[anime_num++].src
			if(anime_ID){
				window.setTimeout(__HMBA.gamesAnime(),delay);
			}
		}
	}

	// PRIVATE FUNCTION
	this.animeSlower = function(){
		if(delay < delay_MAX){
			delay += 15
		}
	}

	// PRIVATE FUNCTION
	this.animeFaster = function(){
		if(delay > 0){
			delay -= 15
		}
	}
	
})
