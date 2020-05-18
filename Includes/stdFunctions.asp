<%

' ********************************************
' * WORLDS OF WAR II STANDARD FUNCTIONS FILE *
' ********************************************

function rtnConName(conCode)
	Dim strDescription

	select case conCode
		case "AA"
			strDescription = "Advance Acadamy"
		case "JF"
			strDescription = "Jeep Factory"
		case "MF"
			strDescription = "Matilda Factory"
		case "SF"
			strDescription = "Sherman Factory"
		case "SPF"
			strDescription = "Spitfire Factory"
		case "HMP"
			strDescription = "Hawlker Production Centre"
		case "LF"
			strDescription = "Lancaster Factory"
		case "HF"
			strDescription = "Halifax Manufacture Plant"
		case "WF"
			strDescription = "Wellington Factory"
		case "BF"
			strDescription = "Bomb Factory"
		case "RF"
			strDescription = "Rocket Factory"
		case "MP"
			strDescription = "Main Port"
		case "DC"
			strDescription = "Defence Control Centre"
		case "RS"
			strDescription = "Radar Station"
		case "SC"
			strDescription = "Spy Training Centre"
		case "SRC"
			strDescription = "Spy Resorce Centre"
		case "MP"
			strDescription = "MP Training Centre"
	end select
	
	rtnConName = strDescription

end function

%>