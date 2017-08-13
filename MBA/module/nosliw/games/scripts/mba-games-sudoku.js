// Copyright© 2003-2006. nOsliw Solutions. All rights reserved.
//**Start Encode**


// http://d4maths.lowtech.org/sudoku.txt
var ID_gameClock = null
var __HSudoku

__H.register(__H.UI.Window.HTA.MBA._Module,"Sudoku","Modules Games Sudoku",function Sudoku(){
	this.funcScroll = __HHTA.onScroll;
	this.txt_solve = "Congratulations! You succesfully solved the Su Doku sudoku in "

	//__HLog.log("# Starting " + document.title) // bug! must use new string instead of ""

	this.initialize = function(idx,xy){
		this.rows = 3
		this.cols = 3
		this.isstarted = false
		this.empty = "##"
		this.isHints = false

		this.iStyleIndex = 0
		this.iLevelIndex = 1
		this.sGridIndex = "3x3"
		this.iSymbolsIndex = 0

		this.v = []
		this.a = []

		this.rowscols = this.rows*this.cols
		this.rowspow = Math.pow(this.rows,2)
		this.colspow = Math.pow(this.cols,2)
	}
	this.initialize()

	this.load = function load(idx,xy){
		try{
			this.setModes(idx,xy)

			this.t1 = games_sud_symbols(this)
			games_sud_div(oDivSudokuSymbols,this.t1)
			this.t1.refresh()

			this.t2 = games_sud_game(this)
			this.t2.cellspacing = 3
			games_sud_div(oDivSudoku1,this.t2)
			this.t2.refresh()
			document.recalc()
			this.isBuilt = true
		}
		catch(ee){
			__HLog.error(ee,this);
			this.isBuilt = false
		}
	}

	this.setModes = function(idx,xy){
		this.setStyle()
		this.setLevel()
		this.setGrid(xy)
		this.setSymbols(idx)
	}

	this.setStyle = function(idx){
		this.iStyleIndex = typeof(idx) == "number" ? idx : this.iStyleIndex
		this.isStyleSudoku = this.iStyleIndex == 0 ? true : false
		this.isStyleSudokuX = this.iStyleIndex == 1 ? true : false
		this.isStyleSquiggle = this.iStyleIndex == 2 ? true : false
		this.isStyleSamurai = this.iStyleIndex == 3 ? true : false
	}

	this.setLevel = function(idx){
		this.iLevelIndex = typeof(idx) == "number"  ? idx : this.iLevelIndex
		this.isLevelEasy = this.iLevelIndex == 0 ? true : false
		this.isLevelMedium = this.iLevelIndex == 1 ? true : false
		this.isLevelHard = this.iLevelIndex == 2 ? true : false
		this.isLevelImpossible = this.iLevelIndex == 3 ? true : false
	}

	this.setGrid = function setGrid(xy){
		try{
			try{
				this.sGridIndex = typeof(xy) == "string" ? xy : this.sGridIndex
				if(this.sGridIndex.match(/([1-9])x([1-9])$/ig)){
					this.rows = RegExp.$1
					this.cols = RegExp.$2
				}
				else throw ""
			}
			catch(e1){
				this.rows = 3
				this.cols = 3
			}
			this.rowscols = this.rows*this.cols
			this.rowspow = Math.pow(this.rows,2)
			this.colspow = Math.pow(this.cols,2)
			if(this.isLevelEasy) this.candidateMIN = Math.floor(0.42*this.rowspow*this.colspow)
			else if(this.isLevelMedium) this.candidateMIN = Math.floor(0.325*this.rowspow*this.colspow)
			else if(this.isLevelHard) this.candidateMIN = Math.floor(0.391*this.rowspow*this.colspow)
			else if(this.isLevelImpossible) this.candidateMIN = Math.floor(0.274*this.rowspow*this.colspow)
			this.candidateTotal = Math.pow(this.rowscols,2)
			return true
		}
		catch(ee){
			__HLog.error(ee,this);
			return false
		}
	}

	this.setSymbols = function setSymbols(idx){
		try{
			this.iSymbolsIndex = typeof(idx) == "number"  ? idx : this.iSymbolsIndex
			this.isSymbolNumber = this.iSymbolsIndex == 0 ? true : false
			this.isSymbolAlphaUpper = this.iSymbolsIndex == 1 ? true : false
			this.isSymbolAlphaLower = this.iSymbolsIndex == 2 ? true : false
			this.isSymbolWebdings = this.iSymbolsIndex == 3 ? true : false
			this.isSymbolWingdings = this.iSymbolsIndex == 4 ? true : false
			this.isSymbolHex = this.iSymbolsIndex == 5 ? true : false

			this.v.empty()
			for(var i = 0; i < this.rowscols; i++){
				if(this.isSymbolNumber) this.v[i] = (i+1) // Number
				else if(this.isSymbolAlphaUpper) this.v[i] = String.fromCharCode(i+65) // Alpha Upper
				else if(this.isSymbolAlphaLower) this.v[i] = (String.fromCharCode(i+65)).toLowerCase() // Alpha Lower
				else if(this.isSymbolWebdings) this.v[i] = String.fromCharCode(i+65) // Webdings
				else if(this.isSymbolWingdings) this.v[i] = String.fromCharCode(i+65) // Wingdings
				else if(this.isSymbolHex) this.v[i] = ((i+1).toString(16)).toUpperCase() // Hexdecimal
			}

			this.a.empty()
			this.a = __HUtil.matrix2D(this.rowspow,this.colspow,this.empty)
		}
		catch(ee){
			__HLog.error(ee,this);
		}
	}

	this.getLastSymbol = function(){
		return this.v[this.rowscols-1]
	}

	this.gameClock = function(){
		if(!this.d_idle) this.d_idle = new Date()
		var d = new Date()
		d = d.getDiff(this.d_idle)
		oFormSudoku.idle.value = d.formatHHMMSS(null,true)

		if((this.candidateTotal-this.candidateCurrent) == 0) setTimeout("__HSudoku.solve()",0)
		else oFormSudoku.action.value = this.candidateCurrent + " cell(s) completed. " + (this.candidateTotal-this.candidateCurrent) + " to go!"
	}

	this.start = function(){
		if(ID_gameClock) this.stop()
		ID_gameClock = setInterval("__HSudoku.gameClock()",1000)
		this.isstarted = true
	}

	this.stop = function(){
		if(ID_gameClock != null){
			clearInterval(ID_gameClock)
			ID_gameClock = null
		}
		this.d_idle = null
		oFormSudoku.idle.value = "00:00:00"
		this.isstarted = false
	}

	this.restart = function restart(){
		try{
			__HLog.log("# Restarting Su DoKu")
			this.stop()
			var oElements = document.getElementsByTagName("input");
			for(var i = 0, iLen = oElements.length; i < iLen; i++){
				if(oElements[i].isSud && !oElements[i].isLock && (oElements[i].name).match(/input_sud_([0-9]{1,2})_([0-9]{1,2})/ig)){
					if(this.a) this.a[RegExp.$1][RegExp.$2] = this.empty
					oElements[i].value = " "
					this.candidateCurrent--
					this.candidateNotFixed--
				}
			}
			this.start()
		}
		catch(ee){
			__HLog.error(ee,this);
		}
	}

	this.clear = function clear(){
		try{
			__HLog.log("# Clearing Su DoKu")
			this.stop()
			this.candidateCurrent = 0
			this.candidateFixed = 0
			this.candidateNotFixed = 0

			var oElements = document.getElementsByTagName("input");
			for(var i = 0, iLen = oElements.length; i < iLen; i++){
				if(oElements[i].isSud && (oElements[i].name).match(/input_sud_([0-9]{1,2})_([0-9]{1,2})/ig)){
					if(this.a) this.a[RegExp.$1][RegExp.$2] = this.empty
					oElements[i].value = " "
					if(oElements[i].isLock) oElements[i].unlock()
				}
			}
			if(this.t1){ // need to check on Init
				var c = this.t1.rows[0].cells
				for(var i = 0, iLen = c.length; i < iLen; i++){
					c(i).bgColor = "#EEEEEE"
				}
			}
			document.recalc()
		}
		catch(ee){
			__HLog.error(ee,this);
		}
	}

	this.gameNew = function gameNew(){
		try{
			if(!this.isBuilt) this.load()
			else {
				this.stop()
				this.setModes()
				this.clear()
			}
			
			this.a = __HUtil.matrix2D(this.rowspow,this.colspow,this.empty)
			__HLog.log("# Generating a new Su DoKu grid")
			var xy = ""
			for(var j = k = 0; this.candidateCurrent <= this.candidateMIN; j++){
				var x = (0).random(this.rowspow)
				var y = (0).random(this.colspow)
				while((k++ % this.rowscols) != 0){ // Run a maximal of X times for a valid candidate
					var vv = (0).random(this.v.length)
					if(this.a[x][y] == this.empty && this.isCandidate(x,y,this.v[vv])){
						this.a[x][y] = this.v[vv]
						var o = __H.byIds("input_sud_" + x + "_" + y)
						//if(!o) alert("input_sud_" + i + "_" + y)
						o.value = this.v[vv]
						o.lock(x,y)
						this.candidateCurrent++
						this.candidateFixed++
						k = 0;
						xy = xy.concat(x+":"+y + " = " + this.v[vv] + "\n")
						break;
					}
				}
			}

			this.t1.refresh()
			this.t2.refresh()
			document.recalc()
			this.start()

			__HLog.log("## Generated a grid with " + this.candidateCurrent + " default candidates")
			//alert(xy)
		}
		catch(ee){
			__HLog.error(ee,this);
		}
	}

	this.solve = function solve(){
		try{
			var bSolve = false
			var t = this.txt_solve + this.idle.formatHHMMSS(false," ")
			for(var i = 0; i < this.a.length; i++){
				for(var j = 0; j < this.a[i].length; j++){
					bSolve = false
					for(var v in this.v){
						if(!this.v.hasOwnProperty(v)) continue
						if(this.a[i][j] == this.v[v]){
							bSolve = true
							break;
						}
					}
					if(!bSolve) break;
				}
				if(!bSolve) break;
			}
			if(bSolve){
				__HLog.popup(t)
				__HLog.log("# " + t)
				this.stop()
			}
			else __HLog.popup("Nope")
		}
		catch(ee){
			__HLog.error(ee,this);
		}
	}

	this.isCandidate = function isCandidate(x,y,v){
		try{
			// 1. Check disctinct horizontal and vertical lines
			__HLog.debug("# (X:Y) = "+x+":"+y + ", f(X:Y) = " + v)
			for(var i = 0; i < this.rowspow; i++){
				__HLog.debug("## Row-"+x+":"+i + " "+this.a[x][i])
				if(this.a[x][i] == v) return false
				else if(this.a[x][i] == this.empty){ // checks for single only cells required for v
					var bSingleHorizontally = true
					for(var m = 0; m < this.colspow; m++){
						if(m != i && this.a[i][m] == this.empty){
							bSingleHorizontally = false
							break;
						}
					}
					if(bSingleHorizontally) return false
				}
			}
			for(var i = 0; i < this.colspow; i++){
				__HLog.debug("## Col-"+i+":"+y + " "+this.a[i][y])
				if(this.a[i][y] == v) return false
				else if(this.a[i][y] == this.empty){ // checks for single only cells required for v
					var bSingleVertically = true
					for(var m = 0; m < this.rowspow; m++){
						if(m != i && this.a[m][i] == this.empty){
							bSingleVertically = false
							break;
						}
					}
					if(bSingleVertically) return false
				}
			}

			// 2. Check the sub BOX
			var o = this.getX0Y0(x,y)
			for(var i = 0; i < this.rows; i++){
				for(var j = 0; j < this.cols; j++){
					if(this.a[o.x0+i][o.y0+j] == v) return false
				}
			}

			// 3. Check last candidate in horizontal and vertival lines for the a.y


			if(this.isDebugisDebug) __HLog.debug("## (X0:Y0) = "+o.x0+":"+o.y0)
			return true
		}
		catch(ee){
			__HLog.error(ee,this);
			return false
		}
	}

	this.getCandidates = function getCandidates(x,y){
		try{
			var a = []
			for(var i = 0; i < this.rowspow; i++){
				if(this.a[x][i] != this.empty) a.push(this.a[x][i])
			}
			for(var i = 0; i < this.colspow; i++){
				if(this.a[i][y] != this.empty) a.push(this.a[i][y])
			}
			var o = this.getX0Y0(x,y)
			for(var i = 0; i < this.rows; i++){
				for(var j = 0; j < this.cols; j++){
					if(this.a[o.x0+i][o.y0+j] != this.empty) a.push(this.a[o.x0+i][o.y0+j])
				}
			}
			return a.sort();
		}
		catch(ee){
			__HLog.error(ee,this);
			return a
		}
	}

	this.getX0Y0 = function(x,y){
		var o = {}
		o.x0 = o.y0 = 0
		for(var i = 0; i < x && i < this.rowspow; i++){
			if(((i+1) % this.rows) == 0 && i < this.rowspow-1){
				o.x0 = i+1
			}
		}
		for(var i = 0; i < y && i < this.colspow; i++){
			if(((i+1) % this.rows) == 0 && i < this.colspow-1){
				o.y0 = i+1
			}
		}
		return o
	}
})


function games_sud_div(o,t){
	try{
		o.style.height = '100%'
		o.style.width = '100%'
		o.style.verticalAlign = 'middle'
		//o.unselectable = "on"
		//o.contentEditable = true
		if(typeof(t) == "object"){
			o.innerText = o.innerHTML = ""
			o.appendChild(t)
		}
	}
	catch(ee){
		__HLog.error(ee,this)
	}
}

function games_sud_cell(o,idx,bVal,x,y){
	try{
		o.noWrap = true
		o.innerText = ""
		o.align = 'center'
		o.style.verticalAlign = 'middle'
		o.isTable = false
		o.isCellChangable = true

		o.onmouseover = function(){
			if(!this.isCellChangable) return;
			this.style.cursor = "hand"
			this.bgColor = __HMBA.cur_color.col_normal
			this.firstChild.focus()
		}
		o.onmouseout = function(){
			if(!this.isCellChangable) return;
			this.style.cursor = "default"
			this.bgColor = ""
			this.style.color = "#000000"
			this.firstChild.style.borderColor = "#000000"
		}
		o.onmousedown = function(){
			if(!this.isCellChangable) return;
			this.bgColor = __HMBA.cur_color.col_dark
			//this.style.color = "#FFFFFF"
			this.firstChild.style.borderColor = "#FFFFFF"
			var n = this.firstChild.name
			if(n.match(/input_sud_([0-9]{1,2})_([0-9]{1,2})/ig)){
				var a = __HSudoku.getCandidates(RegExp.$1,RegExp.$2)
				var c = __HSudoku.t1.rows[0].cells
				for(var i = 0, iLen = c.length; i < iLen; i++){
					c(i).bgColor = "#EEEEEE"
					if(__HSudoku.isHints && this.firstChild.value != " ") continue
					for(var j = 0, iLen2 = a.length; j < iLen2; j++){
						if(c(i).firstChild.value == a[j]){
							c(i).bgColor = "#FFEEEE"
							break;
						}
					}
				}

				//alert(this.x+":"+this.y)
			}
		}
		var f = __H.byClone("input")
		f.xx = o.xx = x
		f.yy = o.yy = y
		f.type = "text"
		f.size = 1
		f.maxlength = 2
		f.className = "cTDInput"
		f.value = typeof(idx) == "number" ? __HSudoku.v[idx] : " "; // must be a space
		//f.unselectable = "on"
		f.isCellChangable = true

		if(__HSudoku.isSymbolWebdings) f.style.fontFamily = "Webdings"
		else if(__HSudoku.isSymbolWingdings) f.style.fontFamily = "Wingdings"
		else f.style.fontFamily = "Arial"

		o.resetCell = function(bLock){
			this.onmouseout()
			//f.style.color = "#0000FF"
			if(bLock) f.isCellChangable = this.isCellChangable = false
		}

		f.onkeydown = function(){
			if(!this.isCellChangable) return false;
			var oEvent = window.event
			var sToChar = String.fromCharCode(oEvent.keyCode); // excellence

			if(oEvent.altKey || oEvent.ctrlKey) return false
			if(__HSudoku.isSymbolNumber && oEvent.keyCode >= 48 && oEvent.keyCode <= 57){ // 0-9
				sToChar = (__HSudoku.rowscols > 9 && this.value.length == 1 && this.value != " ") ? this.value+sToChar : sToChar
				if(sToChar > __HSudoku.getLastSymbol()){
					games_sud_echo("# INFO: Either you entered a too high numeric key '" + sToChar + "' or the key is not within boundaries.")
				}
				else {
					this.value = sToChar
					o.resetCell()
					__HSudoku.candidateCurrent++
					__HSudoku.candidateNotFixed++
				}
			}
			else if((__HSudoku.isSymbolAlphaUpper || __HSudoku.isSymbolAlphaLower || __HSudoku.isSymbolWebdings || __HSudoku.isSymbolWingdings) && oEvent.keyCode >= 65 && oEvent.keyCode <= 90){
				var c = __HSudoku.getLastSymbol()
				if(__HSudoku.isSymbolAlphaLower){
					c = c.toLowerCase(), sToChar = sToChar.toLowerCase()
				}
				if(sToChar <= c){
					this.value = sToChar
					o.resetCell()
					__HSudoku.candidateCurrent++
					__HSudoku.candidateNotFixed++
				}
				else games_sud_echo("# INFO: Either you entered a too high alpha key '" + sToChar + "' or the key is not within boundaries.")
			}
			else if(__HSudoku.isSymbolHex && (oEvent.keyCode >= 48 && oEvent.keyCode <= 57) || (oEvent.keyCode >= 65 && oEvent.keyCode < 71)){
				// Hex is 0-9 and a-f
				if(oEvent.keyCode >= 65 && parseInt(sToChar,16) > __HSudoku.rowscols){ // if alpha
					games_sud_echo("# INFO: Either you entered a too high hex key or the key was not within boundaries.")
				}
				else {
					if(__HSudoku.rowscols > 15 && this.value.length == 1 && this.value != " "){
						var t = this.value+sToChar
						sToChar = (parseInt(t) <= __HSudoku.getLastSymbol()) ? t : " "
					}
					this.value = sToChar
					o.resetCell()
					__HSudoku.candidateCurrent++
					__HSudoku.candidateNotFixed++
				}
			}
			else if(oEvent.keyCode == 32 || oEvent.keyCode == 8){
				// Do this on space, backspace or 0
				this.value = ""
				this.isCellChangable = true
				__HSudoku.candidateCurrent--
				__HSudoku.candidateNotFixed--
			}
			else {
				games_sud_echo("# INFO: You entered an illegal key for the cell (numeric, alpha or hex)")
			}
			event.returnValue = false
		}

		f.onpaste = function(){
			//clipboardData.setData('Text',clipboardData.getData('Text'))

			return true
		}

		f.lock = function(x,y){
			this.readonly = true
			this.isCellChangable = o.isCellChangable = false
			this.unselectable = "on"
			if(!isNaN(x) && !isNaN(y)) this.name = this.id = "input_sud_" + x + "_" + y
			this.isLock = true
			this.style.color = "#000000"
		}

		f.unlock = function(x,y){
			this.readonly = false
			this.isCellChangable = o.isCellChangable = true
			this.unselectable = "yes"
			if(!isNaN(x) && !isNaN(y)) this.name = this.id = "input_sud_" + x + "_" + y
			this.isLock = false
			this.style.color = "#0000FF"
		}
		if(bVal){
			f.lock()
			f.name = f.id = "input_val_" + x + "_" + y
		}
		else {
			f.unlock(x,y)
			f.isSud = true
			//__HSudoku.echo("n = " + f.name)
		}
		o.appendChild(f)
	}
	catch(ee){
		__HLog.error(ee,this)
	}
}

function games_sud_game(s){
	try{
		var t = __H.byClone("table")
		t.width = "100%"
		t.height = "100%"
		t.align = "center"
		t.border = 0
		t.cellspacing = 0
		t.cellpadding = 0
		if(s.rows != s.cols){
			var tt = s.rowspow
			s.rowspow = s.colspow
			s.colspow = tt
		}
		var xy = ""
		for(var j = 0; j < s.colspow; j++){
			var r = t.insertRow(j)
			r.style.verticalAlign = 'top'
			for(var i = 0; i < s.rowspow; i++){
				var c = r.insertCell()
				if(((i+1) % s.cols) == 0 && i < s.colspow-1){
					c.style.borderRightWidth = "4px"
					c.style.borderRightColor = __HMBA.cur_color.col_dark
					c.style.borderRightStyle = "solid"
				}
				if(((j+1) % s.rows) == 0 && j < s.rowspow-1){
					c.style.borderBottomWidth = "4px"
					c.style.borderBottomColor = __HMBA.cur_color.col_dark
					c.style.borderBottomStyle = "solid"
				}
				games_sud_cell(c,false,false,i,j)
				xy = xy.concat(i+":"+j + " ")
			}
			xy = xy.concat("\n")
		}
		if(s.rows != s.cols){
			var tt = s.rowspow
			s.rowspow = s.colspow
			s.colspow = tt
		}
		//alert(xy)
		return t
	}
	catch(ee){
		__HLog.error(ee,this)
		return null
	}
}

function games_sud_symbols(s){
	try{
		var t = __H.byClone("table")
		t.width = "100%"
		//t.height = "30"
		t.align = "center"
		t.border = 0
		t.cellspacing = 3
		t.cellpadding = 0
		var r = t.insertRow(0)
		r.style.verticalAlign = 'top'

		for(var i = 0; i < s.rowscols; i++){
			var c = r.insertCell()
			if(((i+1) % s.rows) == 0 && i < s.rowscols-1){
				c.style.borderRightWidth = "4px"
				c.style.borderRightColor = "#545454"
				c.style.borderRightStyle = "solid"
			}
			games_sud_cell(c,i,true,0,i)
			c.bgColor = "#EEEEEE"
		}

		return t
	}
	catch(ee){
		__HLog.error(ee,this)
		return null
	}
}

function games_sud_reload(){
	try{
		__HSudoku.stop()
		__HSudoku.t1.refresh()
		__HSudoku.t2.refresh()
		oFormSudoku.reset()
		oDivSudoku1.innerHTML = "&nbsp;"
		document.recalc()
		__HSudoku.kill(sud)
	}
	catch(ee){}
}

function games_sud_echo(s){
	try{
		__HSudoku.isEcho = true
		__HLog.log(s)
	}
	catch(ee){
		__HLog.error(ee,this)
	}
}