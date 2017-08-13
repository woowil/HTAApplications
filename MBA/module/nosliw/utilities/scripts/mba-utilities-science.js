// Copyright© 2003-2008. nOsliw Solutions. All rights reserved.
//**Start Encode**


__H.register(__H.UI.Window.HTA.MBA._Module,"Science","Science & Logic",function Science(){
	var oForm = oFormUtilityScience
	var b_initialized = false
	
	this.initialize = function initialize(){
		if(b_initialized) return;
		
		
		b_initialized = true
	}
	
	this.quadSolve = function quadSolve(){
		try{
			var a = parseInt(oForm.Xa.value)
			var b = parseInt(oForm.Xb.value)
			var c = parseInt(oForm.Xc.value)
			var i = 0, str = a+b+c, valid = false
			var X1, X2, tmpquad, tmpquad1, tmpquad2
			
			while(i < str.length){
				if(((ch = str.substring(i,++i)) >= 0 && ch <= 9) || ch == "." || ch == "-") continue;			
				else valid = true; break;
			}
			
			if(!valid && a != 0){ 
				var de = parseFloat("1E+5"), DEC;
				if((DEC = b*b - 4*a*c) < 0){
					tmpquad = Math.sqrt(-DEC), tmpquad1 = Math.ceil((-b/(2*a))*de)/de; tmpquad2 = Math.ceil((tmpquad/(2*a))*de)/de;
					X1 = tmpquad1 + " + " + tmpquad2 + "*I", X2 = tmpquad1 + " - " + tmpquad2 + "*I";	
				}
				else {
					tmpquad = Math.sqrt(DEC);
					tmpquad1 = (-b+tmpquad)/(2*a), X1 = Math.ceil(tmpquad1*de)/de;
					tmpquad2 = (-b-tmpquad)/(2*a), X2 = Math.ceil(tmpquad2*de)/de;
				}
				oForm.X1.value = " " + X1
				oForm.X2.value = " " + X2;
			}
			else {
				oForm.X1.value = oForm.X2.value = " NaN: Not a Number\(s\)";
				__HLog.popup("Some how the value(s) of \"a, b or c\" are incorrect. The statements gives an invalid\nQuadratic Equation. Check so that \"a, b and c\" are real numbers and \"a\" is not zero.");
			}		
			return true
		}
		catch(e){
			__HLog.error(e,this);
			return false
		}
	}

	this.quadClean = function quadClean(){
		try{
			var a = ["Xa","Xb","Xc","X1","X2"]
			for(var i = 0; i < a.length; i++){
				if(a[i] in oForm) oForm[a[i]].value = ""
			}
		}
		catch(e){
			__HLog.error(e,this);
			return false
		}
	}

	var A = B = [];

	this.matrixInsert = function matrixInsert(){
		var m, num
		var val = oForm.matrixval.value
		num = oForm.matrix.options[m = oForm.matrix.options.selectedIndex].value;
		if(m < 9) A[num.substring(1,3)] = parseFloat(val);
		else B[num.substring(1,2)] = parseFloat(val);
		if(oForm.options[++m]){
			oForm.options[m].selected = true
			oForm.matrixval.value = "", oForm.matrixval.focus();
		}
	}

	this.matrixSolve = function matrixSolve(){
		var DET = A[11]*(A[22]*A[33]-A[32]*A[23])-A[12]*(A[21]*A[33]-A[31]*A[23])+A[13]*(A[21]*A[32]-A[31]*A[22]);
		if(DET != 0){
	      	oForm.matrixX1.value = (B[1]*(A[22]*A[33]-A[32]*A[23])-A[12]*(B[2]*A[33]-B[3]*A[23])+A[13]*(B[2]*A[32]-B[3]*A[22]))/DET;
			oForm.matrixX2.value = (A[11]*(B[2]*A[33]-B[3]*A[23])-B[1]*(A[21]*A[33]-A[31]*A[23])+A[13]*(A[21]*B[3]-A[31]*B[2]))/DET;
	 		oForm.matrixX3.value = (A[11]*(A[22]*B[3]-A[32]*B[2])-A[12]*(A[21]*B[3]-A[31]*B[2])+B[1]*(A[21]*A[32]-A[31]*A[22]))/DET;
			if(isNaN(DET)) __HLog.popup('Your values of A and/or B are incomplete or invalid numbers.\nEnter new values for each and every one and try again :)');
		}
		else
			__HLog.popup('The Determinant of A is equal to zero, i.e. det(A)=0, therefore invalid!.\nThat means that your Matrix in inconsistent, please try again!');
		oForm.matrix.options[0].selected = true
		oForm.matrixval.value = 'Det(A)= ' + DET;
	}

})

var __HScience = new __H.UI.Window.HTA.MBA._Module.Science()

