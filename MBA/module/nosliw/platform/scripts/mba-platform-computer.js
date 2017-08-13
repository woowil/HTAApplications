

function platform_computer_chassistype(sComputer,sUser,sPass){
	try{
		var oService
		if(!__HWMI.setServiceWMI(sComputer,"root\\cimv2",sUser,sPass))) return false
		
		var oService = __HWMI.setServiceWMI()
		var oColItems = oService.ExecQuery("Select ChassisTypes from Win32_SystemEnclosure","WQL",48);
		for(var oEnum = new Enumerator(oColItems); !oEnum.atEnd(); oEnum.moveNext()){
			var oItem = oEnum.item();
			var o = {}
			var v = new VBArray(oItem.ChassisTypes).toArray(), v = v[0]
			o.number = v, o.name = "ChassisType"
			switch(parseInt(o.number)){
				case 1 : o.type = "Other"; break;
				case 2 : o.type = "Unknown"; break;
				case 3 : o.type = "Desktop"; break;
				case 4 : o.type = "Low-profile Desktop"; break;
				case 5 : o.type = "Pizza Box"; break;
				case 6 : o.type = "Mini Tower"; break;
				case 7 : o.type = "Tower"; break;
				case 8 : o.type = "Portable"; break;
				case 9 : o.type = "Laptop"; break;
				case 10 : o.type = "Notebook"; break;
				case 11 : o.type = "Hand-held"; break;
				case 12 : o.type = "Docking Station"; break;
				case 13 : o.type = "All-in-one"; break;
				case 14 : o.type = "Sub Notebook"; break;
				case 15 : o.type = "Space-saving"; break;
				case 16 : o.type = "Lunch Box"; break;
				case 17 : o.type = "Main System Chassis"; break;
				case 18 : o.type = "Expansion Chassis"; break;
				case 19 : o.type = "Sub Chassis"; break;
				case 20 : o.type = "Bus-expansion Chassis"; break;
				case 21 : o.type = "Peripherical Chassis"; break;
				case 22 : o.type = "Storage Chassis"; break;
				case 23 : o.type = "Rack-mounted Chassis"; break;
				case 24 : o.type = "Sealed-case Computer"; break;
				default : o.type = "Unknown"; break;
			}
			return o
		}
		return false
	}
	catch(e){
		__HLog.error(e,this)
		return false;
	}
}

function platform_computer_domainrole(sComputer,sUser,sPass){
	try{
		var oService = __HWMI.setServiceWMI(sComputer,"root\\cimv2",sUser,sPass);
		var oColItems = oService.ExecQuery("Select DomainRole from Win32_ComputerSystem","WQL",48);
		for(var oEnum = new Enumerator(oColItems); !oEnum.atEnd(); oEnum.moveNext()){
			var oItem = oEnum.item();
			var o = {}
			o.number = oItem.DomainRole, o.name = "DomainRole"
			switch(parseInt(o.number)){
				case 0 : o.type = "Standalone Workstation"; break;
				case 1 : o.type = "Member Workstation"; break;
				case 2 : o.type = "Standalone Server"; break;
				case 3 : o.type = "Member Server"; break;
				case 4 : o.type = "Backup Domain Controller"; break;
				case 5 : o.type = "Primary Domain Controller"; break;
				default : o.type = "Unknown"; break;
			}
			return o
		}
		return false
	}
	catch(e){
		__HLog.error(e,this)
		return false;
	}
}