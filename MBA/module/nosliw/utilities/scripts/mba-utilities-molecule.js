// Copyright© 2003-2008. nOsliw Solutions. All rights reserved.
//**Start Encode**

__H.register(__H.UI.Window.HTA.MBA._Module,"Molecule","Molecule Formular",function Molecule(){
	var oForm = oFormUtilityMolecule
	var a_atoms
	var d_atoms
	var s_atoms = ""
	var i_atoms_len
	var s_molecule = ""
	var b_initialized = false
	var NOVALUE = parseFloat(0.0000);
	
	this.initialize = function initialize(){
		if(b_initialized) return;
		d_atoms = new ActiveXObject("Scripting.Dictionary")
		
		a_atoms = ["H1.0079","He4.00260","D2.014102","Li6.941","Be9.01218","B10.81","C12.011","N14.0067","O15.9994","F18.99840","Ne20.179","Na22.98977","Mg24.305","Al26.98154","Si28.086","P30.97376","S32.06","Cl35.453","K39.098","Ar39.948","Ca40.08","Sc44.9559","Ti47.90","V50.9414","Cr51.996","Mn54.9380","Fe55.847","Ni58.71","Co58.9332","Cu63.546","Zn65.38","Ga69.72","Ge72.59","As74.9216","Se78.96","Br79.904","Kr83.80","Rb85.4678","Sr87.62","Y88.9059","Zr91.22","Nb92.9064","Mo95.94","Tc99.0","Ru101.07","Rh102.9055","Pd106.4","Ag107.868","Cd112.40","In114.82","Sn118.69","Sb121.75","I126.9054","Te127.60","Xe131.30","Cs132.9054","Ba137.34","La138.9055","Ce140.12","Pr140.9077","Nd144.24","Pm147.0","Sm150.4","Eu151.96","Gd157.25","Tb158.9254","Dv162.50","Ho164.9304","Er167.26","Tm168.9342","Yb173.04","Lu174.97","Hf178.49","Ta180.9479","W183.85","Re186.2","Os190.2","Ir192.22","Pt195.09","Au196.9665","Hg200.59","Tl204.37","Pb207.2","Bi208.9804","At210.0","Po210.0","Rn222.0","Fr223.0","Ra226.0","Ac227.0","Pa231.0","Th232.0381","Np237.0","U238.029","Pu242","Am243.0","Bk247.0","Cm247.0","Cf251.0","Es254.0","No255.0","Lr256.0","Fm257.0","Md257.0","Ha260.0"];
		for(var j = 0, i_atoms_len = a_atoms.length; j < i_atoms_len; j++){
			if(a_atoms[j].match(/^([a-z]{1,2})([0-9.]+)$/ig)){
				if(!d_atoms.Exists(RegExp.$1)){
					d_atoms.Add(RegExp.$1,parseFloat(RegExp.$2))			
				}
			}
		}	
		
		b_initialized = true;
	}	
	
	this.setFormula = function setFormula(){
		this.initialize();
		
	}
	
	this.getFormula = function getFormula(){
		this.setFormula()
		
		s_molecule = oForm.molecule.value;
		oForm.weight.value = this.validate() ? split(s_molecule,0) : NOVALUE;
		oForm.molecule.focus();
		oForm.total.value = parseFloat(oForm.total.value+oForm.weight.value)
	}

	this.validate = function validate(){
		for(var j = i = 0, iLen = s_molecule.length; j <= iLen;  ){
		    var ch = s_molecule.substring(j,++j);
			// ENSURE CHARACTER IS AN ALPHA CHARACTER OR A NUMBER
			if((ch >= "a" && ch <= "z") || (ch >= "A" && ch <= "Z" ) || (ch >= 0 && ch <= 9));
			else if(ch == "(") i++; 
			else if(ch == ")") i--;
			else{
				__HLog.popup("Invalid character: \"" + ch + "\" in your formula");
				return false;
			}		
		}
		if(i != 0){ 
			__HLog.popup("Unbalanced paranthesis in formula:\n" + s_molecule);
			return false;
		}
		return true;
	}

	var atomWeight = function atomWeight(name){
		if(!d_atoms.Exists(name)){
			return d_atoms(name)				
		}
		return NOVALUE;
	}

	var count = function count(s,index){
	    var res1 = 1.0000, res2, ch = '', ok = false;
		
		if(index != s.length && (ch = s.substring(index,++index)) >= 0 && ch <= 9){
			res2 = ch, ok = true;
			while(index != s.length && (ch = s.substring(index,++index)) >= 0 && ch <= 9)
		    	res2 += ch;
		}
		return ok ? res2 : res1;
	}

	var split = function split(s,index){
	    var ss = '', str;
		
	    var ch = s.substring(index,index+1);
		if(index >= s.length) return NOVALUE;
		else if(ch >= "A" && ch <= "Z" ){
			ss = ss.concat(ch)
			while(index != s.length && ((ch = s.substring(index+1,index+2)) >= "a") && ch <= "z"){
		    	ss = ss.concat(ch), index++;
			}
			str = ss.substring(0,ss.length);
			return atomWeight(str) * count(s,++index) + split(s,index);
	    }
		else if(ch == '('){
			for(var i = 1, j = k = 0; i != j && k <= s.length; k++){
		   		if ((ch = s.substring(++index,index+1)) == '(') ++i;
				else if(ch == ')') --i;
				ss = ss.concat(ch)
			}	
			var len = ss.length, ss = ss.substring(0,len-1)
			str = ss.substring(0,ss.length);
			
			return split(str,0) * count(s,++index) + split(s,index);
	    }
	    else return split(s,++index);
	}
})

var __HMolecule = new __H.UI.Window.HTA.MBA._Module.Molecule()
