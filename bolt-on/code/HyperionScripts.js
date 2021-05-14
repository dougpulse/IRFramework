/*
	WSDOT Hyperion Interactive Reporting Script Library
	
	Purpose:	Provide more functionality in Hyperion IR JavaScript to 
				enable developers to code to a set of pre-defined object
				members for standardization, ease of coding, and reusability.
	
	Author:		Douglas Pulse
	
	Comments:	Many of the methods defined herein depend on the existence of
				other methods defined herein.  For example, the String.trim() 
				methods relies on the String.lTrim() and String.rTrim() 
				methods.
				This code will work in Hyperion Interactive Reporting Studio 
				and Hyperion Interactive Reporting Web Client version 9.3.1.0, 
				9.3.1.3, and 11.1.2.  Although it hasn't been tested in them, 
				it may work in other versions as well.  It does not work 
				properly in Hyperion Interactive Reporting iHTML due to its 
				dependency on the JOOLE object.
*/

/*****************************************************************************/
//			Don't forget to point this to the correct file.
//			(dev vs. prod)
/*****************************************************************************/
eval( (new JOOLEObject("Scripting.FileSystemObject")).OpenTextFile(-2, 1, "\\\\hqolymhyperion2\\SharedResources\\Script\\IRFramework.js").ReadAll() )
//eval( (new JOOLEObject("Scripting.FileSystemObject")).OpenTextFile(-2, 1, "C:\\DMSVN\\SharedResources\\Reporting Tool\\Trunk\\IRFramework.js").ReadAll() );

ActiveDocument.logError = new Application.logger((new Application.logger()).LEVEL.ERROR, "console");
ActiveDocument.logTrace = new Application.logger((new Application.logger()).LEVEL.OFF, "console");
ActiveDocument.logProcess = new Application.logger((new Application.logger()).LEVEL.INFO, "console");
Application.LibraryLogger = new Application.logger((new Application.logger()).LEVEL.INFO, "\\\\hqolymhyperion2\\ScriptLogs\\library.log", true);

Number.prototype.pad = function (NumDigits) {  return this.toPaddedString(NumDigits);  }
Number.prototype.zf = Number.prototype.pad;
ActiveDocument.wsdot_ProgressBar = ActiveDocument.ProgressBar;
ActiveDocument.wsdot_StopWatch = Application.StopWatch;
ActiveDocument.wsdot_msToTime = Application.msToTime;
ActiveDocument.wsdot_SynchModels = ActiveDocument.SynchModels;
ActiveDocument.wsdot_SectionExists = ActiveDocument.SectionExists;
ActiveDocument.wsdot_CustomCalendar = Application.CustomCalendar;
ActiveDocument.Font = Application.Font;
ActiveDocument.wsdot_ExternalVBScript = Application.ExternalVBScript;


AddFavorite = function () {
/*
	Object Method:	ActiveDocument.AddFavorite()
	Purpose:	Adds the URL of the ActiveDocument as a favorite in Internet Explorer.
	Returns:	1
	Inputs:		(none)
	
	Revision History
	Date		Developer		Description
	7/15/2010	D. Pulse		inagural
*/
	//	get the smartcut from the workspace
	var strSmartCut = new String(ActiveDocument.URL);
	//	locate the end of the reference for the link
	var iLoc = strSmartCut.indexOf("&repository_name");
	//	extract just the link information
	strSmartCut = strSmartCut.substr(0, iLoc);
	//	prepend the link to the pre-processing file
	//	This file opens the link in an IE window without menus or toolbars
	var strPreCut = "http://datamining/prebqy.htm?";
	//	append the argument to ensure that the file is opened in Insight
	var strPostCut = "?bqtype=insight";
	//	create the link to be included in the message
	var sURL = strPreCut + strSmartCut + strPostCut;
	var sTitle = ActiveDocument.Name;
	
	Application.OpenURL("http://datamining/AddFavorite.htm?url=" + sURL + "&title=" + sTitle, "_new");
	
	return 1;
};
ActiveDocument.wsdot_AddFavorite = AddFavorite;



ActiveDocument.SortMonthNames = function (a, b) {
	//	function to be used as an argument to Array.sort()
	//	sorts the month names in the correct order for the calendar year
	var retVal = 0;
	var x = a.toProperCase();
	var y = b.toProperCase();
	if (!MonthName.contains(x) && !MonthName.contains(y)) {
		if (x == "Month25") {
			retVal = -1;
		}
		else if (y == "Month25") {
			retVal = 1;
		}
		else {
			retVal = ((x == y) ? 0 : ((x < y) ? -1 : 1));
		}
	}
	else if (!MonthName.contains(x)) {
		retVal = 1;
	}
	else if (!MonthName.contains(y)) {
		retVal = -1;
	}
	else {
		retVal = MonthName.indexOf(x) - MonthName.indexOf(y);
	}
	return retVal;
}



ActiveDocument.SortMonthShortNames = function (a, b) {
	//	function to be used as an argument to Array.sort()
	//	sorts the month names in the correct order for the calendar year
	var retVal = 0;
	var x = a.toProperCase().trim();
	var y = b.toProperCase().trim();
	if (!MonthAbbrev.contains(x) && !MonthAbbrev.contains(y)) {
		if (x == "Mo25") {
			retVal = -1;
		}
		else if (y == "Mo25") {
			retVal = 1;
		}
		else {
			retVal = ((x == y) ? 0 : ((x < y) ? -1 : 1));
		}
	}
	else if (!MonthAbbrev.contains(x)) {
		retVal = 1;
	}
	else if (!MonthAbbrev.contains(y)) {
		retVal = -1;
	}
	else {
		retVal = MonthAbbrev.indexOf(x) - MonthAbbrev.indexOf(y);
	}
	return retVal;
}



ActiveDocument.SortFY1MonthNames = function (a, b) {
	//	function to be used as an argument to Array.sort()
	//	sorts the month names in the correct order for FY1
	var retVal = 0;
	var x = a.toProperCase();
	var y = b.toProperCase();
	if (!MonthName.contains(x) && !MonthName.contains(y)) {
		if (x == "Month99") {
			retVal = -1;
		}
		else if (y == "Month99") {
			retVal = 1;
		}
		else {
			retVal = ((x == y) ? 0 : ((x < y) ? -1 : 1));
		}
	}
	else if (!MonthName.contains(x)) {
		retVal = 1;
	}
	else if (!MonthName.contains(y)) {
		retVal = -1;
	}
	else {
		retVal = ((MonthName.indexOf(x) + 6) % 12) - ((MonthName.indexOf(y) + 6) % 12);
	}
	return retVal;
}



ActiveDocument.SortFY2MonthNames = function (a, b) {
	//	function to be used as an argument to Array.sort()
	//	sorts the month names in the correct order for FY2
	var retVal = 0;
	var x = a.toProperCase();
	var y = b.toProperCase();
	if (!MonthName.contains(x) && !MonthName.contains(y)) {
		if (x == "Month25") {
			retVal = -1;
		}
		else if (y == "Month25") {
			retVal = 1;
		}
		else {
			retVal = ((x == y) ? 0 : ((x < y) ? -1 : 1));
		}
	}
	else if (!MonthName.contains(x)) {
		retVal = 1;
	}
	else if (!MonthName.contains(y)) {
		retVal = -1;
	}
	else {
		retVal = ((MonthName.indexOf(x) + 6) % 12) - ((MonthName.indexOf(y) + 6) % 12);
	}
	return retVal;
}



ActiveDocument.CleanDashboard = function (Dashboard, TabWidth, TabOffset, ActiveTab, ResetScripts, DashCount, TabHeight, TabFont) {
//ActiveDocument.CleanDashboard = function (Dashboard, ResetScripts) {
	//	Purpose:	
	
	var iErrLine = 1, iLoop = -1;
	//Console.Writeln(this.Parent.Name);
	
	try {
		iErrLine = 8;	var xBase = 7;
		iErrLine = 11;	var oDash, oShape;
		
		//	create dashboard
		iErrLine = 17;	oDash = Dashboard;
		
		//	place wsdot logo
		iErrLine = 20;	oShape = oDash.Shapes["picLogo"];
		iErrLine = 21;	oShape.Placement.YOffset = 4;
		iErrLine = 22;	oShape.Placement.XOffset = xBase;
		iErrLine = 23;	oShape.Placement.Height = 31;
		iErrLine = 24;	oShape.Placement.Width = 198;
		if (ResetScripts) {
			sTemp = "var sApp = \"C:\\\\Program Files\\\\Internet Explorer\\\\iexplore.exe\";\r\n";
			sTemp += "var sArgs = \"http://wwwi.wsdot.wa.gov\";\r\n";
			sTemp += "Application.Shell(sApp, sArgs);\r\n";
			iErrLine = 29;	oShape.EventScripts["OnClick"].Script = sTemp;
		}
		
		//	create and place report header
		iErrLine = 33;	oShape = oDash.Shapes["lblReportHeader"];
		iErrLine = 34;	oShape.Placement.YOffset = 4;
		iErrLine = 35;	oShape.Placement.XOffset = 265;
		iErrLine = 36;	oShape.Placement.Height = 29;
		iErrLine = 37;	oShape.Placement.Width = 482;
		iErrLine = 38;	oShape.Font.Name = "Arial";
		iErrLine = 39;	oShape.Font.Size = 18;
		iErrLine = 40;	oShape.Font.Style = bqFontStyleBold;				//	2
		iErrLine = 41;	oShape.Font.Effect = bqFontEffectNone;				//	0
		iErrLine = 42;	oShape.Font.Color = 0;
		iErrLine = 43;	oShape.VerticalAlignment = bqAlignMiddle;			//	2
		iErrLine = 44;	oShape.Alignment = bqAlignCenter;					//	1
		
		//	create and place toolbar buttons
		//		export
		try {
			if (ResetScripts) {
				sTemp = "\/\/	export the data from the report to pdf\r\n";
				sTemp += "var d = new Date();\r\n";
				sTemp += "var dfile = new String(\"\");\r\n";
				sTemp += "\r\n";
				sTemp += "dfile = g_strUserTempPath + \"\\\\EVProjInfo-\" + d.format(\"yyyymmdd\") + \".pdf\";\r\n";
				sTemp += "ActiveDocument.Sections[this.Parent.Name.replace(\/ \/gi, \"  \")].Export(dfile, bqExportFormatPDF, true, false);\r\n";
				sTemp += "Shell(g_strWinPath + \"\\\\explorer.exe\", dfile);\r\n";
				iErrLine = 54;	oDash.Shapes["picExport"].EventScripts["OnClick"].Script = sTemp;
			}
		}
		catch (e) {
			//	toolbar button doesn't exist -- move on
		}
		//		edit
		try {
			if (ResetScripts) {
				sTemp = "ActiveDocument.Sections[\"R-Contract Status\"].Activate();\r\n";
				sTemp += "ActiveDocument.ShowCatalog = true;\r\n";
				iErrLine = 68;	oDash.Shapes["picEdit"].EventScripts["OnClick"].Script = sTemp;
			}
		}
		catch (e) {
			//	toolbar button doesn't exist -- move on
		}
		//		e-mail
		try {
			if (ResetScripts) {
				sTemp = "wsdot_createEmail(\"\", \"Cost Performance Report\");\r\n";
				iErrLine = 78;	oDash.Shapes["picSend"].EventScripts["OnClick"].Script = sTemp;
			}
		}
		catch (e) {
			//	toolbar button doesn't exist -- move on
		}
		//		print preview
		try {
			if (ResetScripts) {
				sTemp = "ActiveDocument.Sections[this.Parent.Name.replace(\/ \/gi, \"  \")].Activate();\r\n";
				iErrLine = 88;	oDash.Shapes["picPrintPreview"].EventScripts["OnClick"].Script = sTemp;
			}
		}
		catch (e) {
			//	toolbar button doesn't exist -- move on
		}
		//		print
		try {
			if (ResetScripts) {
				sTemp = "ActiveDocument.Sections[this.Parent.Name.replace(\/ \/gi, \"  \")].PrintOut();\r\n";
				iErrLine = 98;	oDash.Shapes["picPrint"].EventScripts["OnClick"].Script = sTemp;
			}
		}
		catch (e) {
			//	toolbar button doesn't exist -- move on
		}
		//		help
		try {
			if (ResetScripts) {
				sTemp = "\r\n";
				sTemp += "\r\n";
				iErrLine = 109;	oDash.Shapes["picHelpD"].EventScripts["OnClick"].Script = sTemp;
			}
		}
		catch (e) {
			//	toolbar button doesn't exist -- move on
		}
		//		reset placement
		wsdot_GlobalStartup.setToolbarButtons();
	}
	catch (e) {
		Console.Writeln("Error on line " + iErrLine + ", loop " + iLoop + ":  " + e.toString());
	}
	ActiveDocument.SetTabBar(Dashboard, TabWidth, TabOffset, ActiveTab, ResetScripts, DashCount, TabHeight, TabFont, "", 47);
}

ActiveDocument.SetTabBar = function (Dashboard, TabWidth, TabOffset, ActiveTab, ResetScripts, DashCount, TabHeight, TabFont, BarSuffix, Top) {
	//	Purpose:	
	
	try {
		if (!TabFont.isFont()) TabFont = new ActiveDocument.Font();
	}
	catch (e) {
		TabFont = new ActiveDocument.Font();
	}
	
	if (typeof BarSuffix == "string") {
		_barSuffix = BarSuffix;
	}
	else {
		_barSuffix = "";
	}
	
	var iErrLine = 1, iLoop = -1;
	//Console.Writeln(this.Parent.Name);
	
	try {
		iErrLine = 6;	var iDash = ActiveTab;
		iErrLine = 7;	var yBase = Top;
		iErrLine = 8;	var xBase = 7;
		iErrLine = 9;	var iWidth = TabWidth;
		iErrLine = 10;	var iHeight = TabHeight;
		iErrLine = 11;	var oDash, oShape;
		iErrLine = 12;	var iOffset = TabOffset;
		iErrLine = 13;	var arrTabLeft = new Array(DashCount);
		iErrLine = 14;	var arrTabWidth = new Array(DashCount);
		
		//	create dashboard
		iErrLine = 17;	oDash = Dashboard;
		//Console.Writeln(oDash.Name);
		
		//	create and place horizontal bars
		iErrLine = 120;	oShape = oDash.Shapes["picBarTop" + _barSuffix];
		iErrLine = 121;	oShape.Placement.YOffset = yBase;
		iErrLine = 122;	oShape.Placement.XOffset = xBase;
		iErrLine = 123;	oShape.Placement.Height = iHeight + 4;
		iErrLine = 124;	oShape.Placement.Width = 978;
		
		iErrLine = 126;	oShape = oDash.Shapes["picBarBottom" + _barSuffix];
		iErrLine = 127;	oShape.Placement.YOffset = yBase + oDash.Shapes["picBarTop" + _barSuffix].Placement.Height;
		iErrLine = 128;	oShape.Placement.XOffset = xBase;
		iErrLine = 129;	oShape.Placement.Height = 10;
		iErrLine = 130;	oShape.Placement.Width = oDash.Shapes["picBarTop" + _barSuffix].Placement.Width;
		
		//	create and place tab text and dividers
		arrTabWidth[0] = iWidth + 1;
		while (arrTabWidth.max() > iWidth) {
			iWidth = arrTabWidth.max();
			iErrLine = 133; for (var i = 0; i < DashCount; i++) {
				iLoop = i;  //Console.Writeln("loop:  " + iLoop);
				iErrLine = 135;	arrTabWidth[i] = iWidth;
				iErrLine = 136;	arrTabLeft[i] = (i == 0 ? xBase + iOffset : arrTabLeft[i - 1] + arrTabWidth[i - 1]);
				
				iErrLine = 138; oShape = oDash.Shapes["lblDash" + _barSuffix + (i + 1).toString()];
				iErrLine = 139; oShape.Placement.YOffset = yBase + 3;
				iErrLine = 140; oShape.Placement.XOffset = arrTabLeft[i];
				iErrLine = 141;	oShape.Placement.Width = arrTabWidth[i] - 13;
				iErrLine = 142; oShape.Placement.Height = iHeight;
				iErrLine = 143; oShape.Font.Name = TabFont.Name();
				iErrLine = 144; oShape.Font.Size = TabFont.Size();
				iErrLine = 145; oShape.Font.Effect = TabFont.Effect();
				iErrLine = 146; oShape.Font.Style = TabFont.Style();
				
				iErrLine = 147; while (oShape.Placement.Height > iHeight) {
					iErrLine = 148;	arrTabWidth[i] += 4;
					iErrLine = 149;	oShape.Placement.Width = arrTabWidth[i] - 13;
					iErrLine = 150; oShape.Placement.Height = iHeight;
				}
				
				if (i == iDash) {
					oShape.Font.Color = 0;
				}
				else {
					oShape.Font.Color = 16777215;
				}
				if (ResetScripts) {
					iErrLine = 156;	oShape.EventScripts["OnClick"].Script = "ActiveDocument.Sections[this.Text].Activate();\r\n";
				}
				iErrLine = 159;	if (i == 0) continue;
				iErrLine = 160; oShape = oDash.Shapes["picBarDivider" + _barSuffix + i.toString()];
				iErrLine = 161; oShape.Placement.YOffset = yBase;
				iErrLine = 162; oShape.Placement.XOffset = arrTabLeft[i] - 7;
				iErrLine = 163; oShape.Placement.Height = iHeight + 4;
			}
		}
		
		//	create and place active tab graphic
		iErrLine = 166; oShape = oDash.Shapes["picTabLeft" + _barSuffix];
		iErrLine = 167;	oShape.Placement.YOffset = yBase;
		iErrLine = 168;	oShape.Placement.XOffset = arrTabLeft[iDash] - 4;
		iErrLine = 173;	oShape.Placement.Height = iHeight + 4;
		
		iErrLine = 169; oShape = oDash.Shapes["picTabMid" + _barSuffix];
		iErrLine = 171;	oShape.Placement.YOffset = yBase;
		iErrLine = 172;	oShape.Placement.XOffset = arrTabLeft[iDash] + 3;
		iErrLine = 173;	oShape.Placement.Width = arrTabWidth[iDash] - 20;
		iErrLine = 173;	oShape.Placement.Height = iHeight + 4;
		
		iErrLine = 174; oShape = oDash.Shapes["picTabRight" + _barSuffix];
		iErrLine = 176;	oShape.Placement.YOffset = yBase;
		iErrLine = 177;	oShape.Placement.XOffset = arrTabLeft[iDash] + arrTabWidth[iDash] - 17;
		iErrLine = 173;	oShape.Placement.Height = iHeight + 4;
		
	//	//	create and place main divider bar
	//	iErrLine = 180;	oShape = oDash.Shapes["linOptions" + _barSuffix];
	//	iErrLine = 181;	oShape.Placement.YOffset = yBase + oDash.Shapes["picBarTop" + _barSuffix].Placement.Height + oDash.Shapes["picBarBottom" + _barSuffix].Placement.Height + 6;
	//	iErrLine = 182;	oShape.Placement.XOffset = arrTabLeft[1] - 7;
	//	iErrLine = 183;	oShape.Placement.Height = 502;
	//	iErrLine = 184;	oShape.Placement.Width = 1;
	}
	catch (e) {
		Console.Writeln("Error on line " + iErrLine + ", loop " + iLoop + ":  " + e.toString());
	}
}



KeyValuePair = function(key, value) {
/*****************************************************************************/
//		Deprecated:  	This class is used only by the Dictionary class.
//						See Dictionary.
/*****************************************************************************/
	Application.LibraryLogger.info("KeyValuePair class is deprecated.");
	var _key = "";
	var _value = "";
	if (typeof key != "number" && typeof key != "string") {
		//	key must be a number or String
	}
	else {
		_key = key;
	}
	if (typeof value == "function") {
		//	value must not be a function
	}
	else {
		_value = value;
	};
	this.key = function(sKey) {
		if (typeof sKey == "undefined") {
			//	do nothing
		}
		else if (typeof sKey == "number" || typeof sKey == "string") {
			_key = sKey;
		}
		return _key;
	};
	this.value = function(vVal) {
		if (typeof vVal == "undefined") {
			//	do nothing
		}
		else if (typeof vVal != "function") {
			_value = vVal;
		}
		return _value;
	};
};
ActiveDocument.wsdot_KVPair = KeyValuePair;


Dictionary = function () {
/*****************************************************************************/
//		Deprecated:  	This class is being replaced by Hash.
//						See Dictionary.
/*****************************************************************************/
/*
	Name:		Dictionary
	Purpose:	Provides a dictionary object for javascript without relying 
				upon Microsoft Windows.
	Inputs: 	(no constructor)
	Returns:	A Dictionary object
	
	Revision History
	Date		Developer	Description
				D. Pulse	inagural
	2/28/2008	D. Pulse	Limited input types for the Add method (which is used 
							by the Key method) to plug a potential security problem.
	4/6/2009	D. Pulse	Change the storage method from object attributes to an array of key/value pairs
	4/2/2010	D. Pulse	Added the Dump() and toString() methods.
*/
	Application.LibraryLogger.info("Dictionary class is deprecated.");
	var arrPairs = new Array();
	
	function PairLocation(vKey) {
		var retVal = -1;
		for (var i = 0; i < arrPairs.length; i++) {
			if (arrPairs[i].key() == vKey) {
				retVal = i;
			}
		}
		return retVal;
	};
	
	this.Count = function () {
		return arrPairs.length;
	};
	this.Add = function (oPair) {
	//this.Add = function (key, value) {
		//	add an item to the dictionary
		var retVal;
		if (arguments.length == 1 && KeyValuePair.prototype.isPrototypeOf(arguments[0])) {
			if (this.Exists(oPair.key())) {
				retVal = -1;	//	key already exists;
				//Console.Writeln(arguments.length.toString() + " arg, " + oPair.key() + " exists");
			}
			else {
				arrPairs.push(oPair);
				retVal = 0;
			}
		}
		else if (arguments.length == 2) {
			var kvTemp = new KeyValuePair(arguments[0], arguments[1]);
			if (this.Exists(kvTemp.key())) {
				retVal = -1;	//	key already exists;
				//Console.Writeln(arguments.length.toString() + " args, key exists");
			}
			else {
				arrPairs.push(kvTemp);
				retVal = 0;
			}
		}
		else {
			retval = -1;	//	incorrect arguments
			//Console.Writeln(arguments.length.toString() + " args");
		}
		return retVal;
	};
	this.Remove = function (vKeyName) {
		//remove an item from the dictionary
		var retVal = -1;
		if (this.Exists(vKeyName)) {
			retVal = 0;
			arrPairs.splice(PairLocation(vKeyName), 1);
		}
		return retVal;
	};
	this.RemoveAll = function () {
		//remove all items from the dictionary
		arrPairs.splice(0, arrPairs.length);
		return 0;
	};
	this.Item = function (vKeyName, vNewValue) {
		//get the value of an item from the dictionary
		var retVal;
		if (this.Exists(vKeyName)) {
			retVal = arrPairs[PairLocation(vKeyName)].value(vNewValue);
		}
		return retVal;
	};
	this.Items = function () {
		//get the collection of all items from the dictionary
		var arrTemp = new Array();
		for (var i = 0; i < arrPairs.length; i++) {
			arrTemp.push(arrPairs[i].value());
		}
		return(arrTemp);
	};
	this.Key = function (vKeyName, vNewName) {
		//get the value of a key from the dictionary
		if (this.Exists(vKeyName)) {
			if (typeof vNewName != "undefined") 
				retVal = arrPairs[PairLocation(vKeyName)].key(vNewName);
		}
		return retVal;
	};
	this.Keys = function () {
		var arrTemp = new Array();
		for (var i = 0; i < arrPairs.length; i++) {
			arrTemp.push(arrPairs[i].key());
		}
		return(arrTemp);
	};
	this.Exists = function (vKeyName) {
		if (PairLocation(vKeyName) == -1) {
			return false;
		}
		else {
			return true;
		}
	};
	this.Promote = function (vKey) {
		arrPairs.promote(PairLocation(vKey));
	};
	this.Demote = function (vKey) {
		arrPairs.demote(PairLocation(vKey));
	};
	this.Dump = function () {
		var arrTemp = new Array();
		arrTemp.push("Count:  " + arrPairs.length);
		for (var i = 0; i < arrPairs.length; i++) {
			arrTemp.push("Key:\r\n" + arrPairs[i].key().toString() + "\r\nValue:\r\n" + arrPairs[i].value().toString());
		}
		return arrTemp;
	};
	this.toString = function() {
		return this.Dump();
	}
};
ActiveDocument.wsdot_Dictionary = Dictionary;



ActiveDocument.wsdot_ClarifySQL = function (SQLStatement) {
/*
	Function:	ClarifySQL()
	
	Type:		JavaScript 
	Purpose:	Reformats a SQL string to be more readable.
	Inputs:		SQLStatement	Required	the SQL string to reformat
	
	Returns:	A string containing the formatted date.
	Notes:		This was designed for use with Hyperion queries, but may 
				work well with other queries.
	
	Revision History
	Date		Developer	Description
	
*/
	//var strTemp = ActiveDocument.Sections["Q-ASOP"].LastSQLStatement;
	var strTemp = SQLStatement;
	
	var reFrom = new RegExp(" FROM ", "gi");
	var reWhere = new RegExp(" WHERE ", "gi");
	var reGroup = new RegExp(" GROUP BY ", "gi");
	var reHaving = new RegExp(" HAVING ", "gi");
	var reOrder = new RegExp(" ORDER BY ", "gi");
	var reAnd = new RegExp(" AND ", "gi");
	var reEq = new RegExp("\S=\S", "gi");
	var reComma = new RegExp(",", "gi");
	
	strTemp = strTemp.replace(reFrom, "\r\n\r\nFROM ");
	strTemp = strTemp.replace(reWhere, "\r\n\r\nWHERE ");
	strTemp = strTemp.replace(reGroup, "\r\n\r\nGROUP BY ");
	strTemp = strTemp.replace(reHaving, "\r\n\r\nHAVING ");
	strTemp = strTemp.replace(reOrder, "\r\n\r\nORDER BY ");
	strTemp = strTemp.replace(reAnd, "\r\n  and ");
	strTemp = strTemp.replace(reEq, " = ");
	strTemp = strTemp.replace(reComma, "\r\n,");
	
	return(strTemp);
};



ActiveDocument.wsdot_ResetFile = function (arrExclude) {
/*
	Function:	ResetFile()
	
	Type:		JavaScript 
	Purpose:	Makes a bqy file smaller for publication by clearing out Results 
				and control values.
	Inputs:		arrExclude	Optional	an array of strings representing the names 
										of Queries or Dashboard Controls to ignore
	
	Returns:	(none)
	
	Revision History
	Date		Developer	Description
	
*/
	var strSection = new String();
	var strTopic = new String();
	var strTopicItem = new String();
	var strTopicItemName = new String();
	var blnClear = new Boolean();
	var blnFilterSet = new Boolean();
	
	for (i = 1; i <= ActiveDocument.Sections.Count; i++) {
		//	the section is a query -- remove all data from the results section
		varSection = ActiveDocument.Sections[i];
		strSection = varSection.Name;
		if (varSection.Type == bqQuery) {
			//	the section is a query -- proceed
			Console.Write("Query:  " + strSection);
			//	if the section name is found in arrExclude, don't clear it
			blnClear = true;
			for (a in arrExclude) {
				if (strSection == arrExclude[a]) {
					blnClear = false;
					Console.Writeln("  --  not cleared (excepted)");
					break;
				}
			}
			blnFilterSet = false;
			if (blnClear) {
				//	we didn't find the section in the exclusion list
				Console.Writeln("  --  clearing");
				varTopic = varSection.DataModel.Topics[1];
				strTopic = varTopic.Name;
				var myLimit;
				for (j = 1; j <= varTopic.TopicItems.Count; j++) {
					strTopicItem = varTopic.TopicItems[j].PhysicalName;
					strTopicItemName = varTopic.TopicItems[j].DisplayName;
					
					if (strTopicItem.toLowerCase().substr(strTopicItem.length - 2) == "id") {
						myLimit = varSection.Limits.CreateLimit(strTopic + "." + strTopicItemName);
						myLimit.CustomValues.Add(-1);
						myLimit.SelectedValues.Add(-1);
						varSection.Limits.Add(myLimit);
						blnFilterSet = true;
						break;
					}
				}
				if (blnFilterSet) {
					varSection.Process();
					varSection.Limits[varSection.Limits.Count].Remove();
					Console.Writeln("Query:  " + strSection + "  --  cleared");
				}
				else {
					Console.Writeln("Query:  " + strSection + "  -- not cleared (no \"id\" field found)");
				}
			}
		}
		if (ActiveDocument.Sections[i].Type == bqDashboard) {
			//	the section is a dashboard -- clear the control values
			Console.Writeln("Dashboard:  " + strSection);
			for (j = 1; j < varSection.Shapes.Count; j++) {
				//	if the control name is found in arrExclude, don't remove it
				varControl = varSection.Shapes[j]
				strControl = varControl.Name;
				Console.Write("Control:  " + strControl);
				blnClear = true;
				for (a in arrExclude) {
					if (strControl == arrExclude[a]) {
						blnClear = false;
						Console.Writeln("  --  not cleared (excepted)");
						break;
					}
				}
				if (blnClear) {
					//	we didn't find the control in the exclusion list
					if (varControl.Type == bqDropDown || varControl.Type == bqListBox) {
						varControl.RemoveAll();
						Console.Writeln("  -  cleared.");
					}
					else if (varControl.Type == bqTextBox) {
						varControl.Text = "";
						Console.Writeln("  -  cleared.");
					}
					else if (varControl.Type == bqCheckBox) {
						varControl.Checked = false;
						Console.Writeln("  --  cleared.");
					}
					else {
						Console.Writeln(" is not of a type that's cleared.");
					}
				}
			}
		}
	}
	for (i = 1; i <= ActiveDocument.Sections.Count; i++) {
		if (ActiveDocument.Sections[i].Type == bqPivot) {
			//	the section is a pivot -- if it is manually refreshed, update the data
			varSection = ActiveDocument.Sections[i];
			if (varSection.RefreshData == bqRefreshDataManually) {
				varSection.RefreshDataNow();
			}
		}
	}
	
	//	remove all data from variable controls on the dashboard
	return(true);
};



ActiveDocument.wsdot_SetLimitFromResult = function (strSourceSectionName, strSourceColName, strSectionName, strItem, strTopic) {
/*
	Function:	setLimitFromResult()
	
	Type:		JavaScript
	Purpose:	Populates a query limit with items in Brio result set.
	Inputs:		strSectionName - The name of the Brio Query section containing the limit.
				strTopic - The topic in the query to use in limit
				strItem - The item in the given topic to use in limit.
				strSourceSectionName - Name of results or table object to get values from.
				strSourceColName - Name of column in result set to get values from.
	
	Returns: None
	Modification History:
	Date		Developer		Description
	03/18/2008	Marj Shomshor	Original
	04/23/2008	Doug Pulse		Added error handling
	02/08/2010	Doug Pulse		Cleaned up documentation and added the ability to use a Table object as the source
*/
	var sMsg, e;
	var varLimit, varSource, varCol, varSctn, strColName, varItem, varTopic;
	
	//	validate input
	//	Is strSourceSectionName a results or table object?
	sMsg = strSourceSectionName + " is not the name of a Results or Table object.";
	try {
		varSource = ActiveDocument.Sections[strSourceSectionName];
	}
	catch (e) {
		//	not a section
		Console.Writeln(sMsg);
		return(sMsg);
	}
	if (varSource.Type != bqResults && varSource.Type != bqTable) {
		//	not a results or table object
		Console.Writeln(sMsg);
		return(sMsg);
	}
	
	//	Is strSourceColName a column in results?
	sMsg = strSourceColName + " is not the name of a column in " + strSourceSectionName + ".";
	try {
		varCol = varSource.Columns[strSourceColName];
	}
	catch (e){
		//	not a control in strSourceSectionName
		Console.Writeln(sMsg);
		return(sMsg);
	}
	
	//	Is strSectionName a query, results, or table?
	sMsg = strSectionName + " is not the name of a Query, Results, or Table section in the ActiveDocument.";
	try {
		varSctn = ActiveDocument.Sections[strSectionName];
	}
	catch (e){
		//	not a section
		Console.Writeln(sMsg);
		return(sMsg);
	}
	if (varSctn.Type != bqQuery && varSctn.Type != bqResults && varSctn.Type != bqTable) {
		//	not a query, results, or table
		Console.Writeln(sMsg);
		return(sMsg);
	}
	
	if (varSctn.Type == bqQuery) {
		//	strSectionName is a query
		//	Is strTopic in strSection?
		try {
			varTopic = varSctn.DataModel.Topics[strTopic];
		}
		catch (e){
			//	not a topic
			sMsg = strTopic + " is not the name of a Topic in " + strSectionName + ".";
			Console.Writeln(sMsg);
			return(sMsg);
		}
		//	Is strItem in strTopic?
		try {
			varItem = varTopic.TopicItems[strItem];
		}
		catch (e){
			//	not an item
			sMsg = strItem + " is not the name of a column in " + strSectionName + ".";
			Console.Writeln(sMsg);
			return(sMsg);
		}
	}
	else {
		//	strSectionName is a results
		//	Is strItem in strSectionName?
		try {
			varItem = varSctn.Columns[strItem];
		}
		catch (e){
			//	not an item
			sMsg = strItem + " is not the name of a column in " + strSectionName + ".";
			Console.Writeln(sMsg);
			return(sMsg);
		}
	}
	
	//	input is good -- continue
	var intItemCount = varSource.RowCount;
	var objColumn  = varCol;

	var varLimit, varCellValue;

	varLimit = varSctn.Limits.CreateLimit(strTopic + "." + strItem);
	varLimit.Name = strItem;
	varLimit.Operator = bqLimitOperatorEqual;

	for(i = 1; i <=  intItemCount; i++)
	{
		varCellValue = varCol.GetCell(i);
		varLimit.CustomValues.Add(varCellValue);
		varLimit.SelectedValues.Add(varCellValue);
	}
	varSctn.Limits.Add(varLimit);
};



ActiveDocument.wsdot_getSelRadioButton = function (strEISSection, strRadioGroup) {
/*
	Function:	getSelRadioButton()
	
	Type:		JavaScript 
	Purpose:	Gets the name of the selected radio button in a group.
	Inputs:		strEISSection	The name of the Brio EIS section containing the list box control.
	 			strRadioGroup	Name of the group.
	
	Returns:	string			name of the selected radio button, zero-length if no button is selected
				string			error message
	Modification History:
	Date		Developer		Description
	04/15/2008	Doug Pulse		Original
*/
	var sMsg, e;
	var sFctn = "wsdot_getSelRadioButton:  ";
	
	//	check for input errors
	//	strEISSection
	sMsg = sFctn + strEISSection + " is not the name of a Dashboard object.";
	try {
		var varDash = ActiveDocument.Sections[strEISSection];
	}
	catch (e) {
		//	doesn't exist
		Console.Writeln(sMsg);
		return(sMsg);
	}
	
	if (varDash.Type != bqDashboard) {
		//	not a dashboard
		Console.Writeln(sMsg);
		return(sMsg);
	}
	
	var NumControls = varDash.Shapes.Count
	var blnGroupFound = false;
	
	for (i = 1; i <= NumControls; i++)
	{
		if (varDash.Shapes[i].Group == strRadioGroup) {
			blnGroupFound = true;
			if (varDash.Controls[i].Checked == true)
				return(varDash.Controls[i].Name);
		}
	}
	if (blnGroupFound) {
		return(false);		//	no options are selected
	}
	else {
		Console.Writeln(strRadioGroup + " is not a control group.");
		return(strRadioGroup + " is not a control group.")
	}
};



ActiveDocument.wsdot_createEmail = function (arrTo, strSubject, strMessage, arrAttach, blnIncludeLink) {
/*
	Function:	wsdot_createEmail()
	
	Modification History:                                             
	Date		Developer	Description
	7/25/2008	Doug Pulse	Original
	1/4/2012	Doug Pulse	Added blnIncludeLink
	8/1/2012	Doug Pulse	DEPRECATED:  Please use ActiveDocument.createEmail
*/
	Application.LibraryLogger.info("wsdot_createEmail class is deprecated.");
	ActiveDocument.createEmail(arrTo, strSubject, strMessage, arrAttach, blnIncludeLink);
};



ActiveDocument.wsdot_ControlExists = function (strDashboardName, strControlName) {
/*****************************************************************************/
//		Deprecated:  	This function is being replaced by the Dashboard.ControlExists() method.
/*****************************************************************************/
/*
	Function:	wsdot_SectionExists()
	
	Type:		JavaScript 
	Purpose:	Determine whether the named section exists in the ActiveDocument.
				
	Inputs:		strControlName		string	name of the control for which to search
				strDashboardName	string	name of the dashboard to search in
	
	Returns:	boolean
	
	Modification History:                                             
	Date		Developer	Description
	10/29/2008	Doug Pulse	Original
*/
	Application.LibraryLogger.info("wsdot_ControlExists function is deprecated.");
	var strTemp = "";
	var intType = -1;
	var blnRetVal = false;
	try {
		//	verify that strDashboardName is the name of a dashboard object
		intType = ActiveDocument.Sections[strDashboardName].Type;
		if (intType == bqDashboard) {
			//	verify that strControlName is the name of a control in strDashboardName
			try {
				strTemp = ActiveDocument.Sections[strDashboardName].Shapes[strControlName].Name;
			}
			catch (e) {
				Console.Writeln(strControlName + " is not a control in " + strDashboardName + ".");
			}
		}
		else {
			Console.Writeln(strDashboardName + " is not a Dashboard object.");
		}
		
	}
	catch (e) {
		Console.Writeln(strDashboardName + " is not a Dashboard object.");
	}
	if (strTemp == "") {
		blnRetVal = false;
	}
	else {
		blnRetVal = true;
	}
	return (blnRetVal);
}



ActiveDocument.wsdot_CheckInputControl = function (strEISSection, strControlName) {
/*****************************************************************************/
//		Deprecated:  	This function is being replaced by the Shape.GetValue method.
/*****************************************************************************/
/*
	Function:	checkInputControl()
	
	Type:		JavaScript 
	Purpose:	Examines a given input control to determine if there is an input value.
	Inputs:		strEISSection	The name of the Brio EIS section containing the list box or textbox control.
				strControlName	The name of the input control to check.
	
	Returns:	boolean		true if a value has been input
							false if no value has been input
	
	Modification History:
	Date		Developer	Description
	5/2/2003	Rich Neill	Original
	3/13/2008	Doug Pulse	Added handling for controls other than ListBox and DropDown
	3/27/2008	Doug Pulse	Added error handling
	9/9/2008	Doug Pulse	Corrected an error with ListBox and DropDown controls
							Since a DropDown always has an item selected, it has been removed.
*/
	Application.LibraryLogger.info("wsdot_CheckInputControl function is deprecated.");
	var varSelectedIndex = new Number();
	var varSelectedValue = new String();
	var varDashboard, varControl;
	var sFctn = "wsdot_CheckInputControl:  ";
	
	var varControl = ActiveDocument.Sections[strEISSection].Shapes[strControlName];
	//	validate the dashboard name
	try {
		varDashboard = ActiveDocument.Sections[strEISSection]
	}
	catch (e) {
		//	strEISSection is not the name of a section
		Console.Writeln(sFctn + strEISSection + " is not the name of a Dashboard object.");
		return false;
	}
	if (varDashboard.Type != bqDashboard) {
		//	strEISSection is the name of a section, but not a dashboard
		Console.Writeln(sFctn + strEISSection + " is not the name of a Dashboard object.");
		return false;
	}
	
	//	validate the control name
	try {
		varControl = varDashboard.Shapes[strControlName];
	}
	catch (e) {
		//	strControlName is not the name of a control
		Console.Writeln(sFctn + strControlName + " is not the name of a control in " + strEISSection + ".");
		return false;
	}
	
	//	check the control type and determine the value of the entered or selected input
	if (varControl.Type == bqListBox) {
		if (varControl.SelectedList.Count == 0) {
			return(false);
		}
		else {
			return(true);
		}
	}
	else if (varControl.Type == bqTextBox) {
		varSelectedValue = varControl.Text;
		if (varSelectedValue.length == 0) {
			return(false);
		}
		else {
			return(true);
		}
	}
	else {
		//	strControlName is not the name of a ListBox or TextBox control
		Console.Writeln(sFctn + strControlName + " is not a ListBox or TextBox control.");
		return false;
	}
};



ActiveDocument.wsdot_GetControlValues = function (strEISSection, strControlName) {
/*****************************************************************************/
//		Deprecated:  	This function is being replaced by the Shape.GetValue method.
/*****************************************************************************/
/*
	Function:	getControlValues()
	
	Type:		JavaScript 
	Purpose:	Collects values from a given control, including Listbox, dropdown box, and radio buttons.
	Inputs:		strEISSection	The name of the Brio EIS section containing the control.
				strListBoxName	Name of the control to evaluate.
	
	Returns:	None
	Modification History:
	Date		Developer		Description
	5/1/2003	Rich Neill		Original
	4/3/2008	Doug Pulse		Added error handling
*/
	Application.LibraryLogger.info("wsdot_GetControlValues function is deprecated.");
	var sMsg, e;
	var sFctn = "wsdot_GetControlValues:  ";
	
	var varControlType;
	var strControlValues;
	var strSelectedItemValues = "";
	var intSelectedCount;
	
	//	check for input errors
	//	strEISSection
	sMsg = sFctn + strEISSection + " is not the name of a Dashboard object.";
	try {
		var varDash = ActiveDocument.Sections[strEISSection]
	}
	catch (e) {
		//	doesn't exist
		Console.Writeln(sMsg);
		return(sMsg);
	}
	
	if (varDash.Type != bqDashboard) {
		//	not a dashboard
		Console.Writeln(sMsg);
		return(sMsg);
	}
	
	//	strControlName
	sMsg = sFctn + strControlName + " is not the name of a control in " + strEISSection + ".";	
	try {
		varControl = varDash.Shapes[strControlName];
		varControlType = varControl.Type;		
		//	verify that the control is of a type that is checked, and get the value.
		switch(varControlType) {
			case bqDropDown :	
				strSelectedItemValues = varControl.Item(varControl.SelectedIndex);
				break;
				
			case bqListBox :
				//	Determine number of items selected in ListBox.
				var intSelectedCount = varControl.SelectedList.Count;
				//	return a space-delimited list of values
				for(i = 1; i <= intSelectedCount; i++) {
					strSelectedItemValues = strSelectedItemValues + varControl.SelectedList.Item(i) + " ";
				}
				break;
				
			case bqRadioButton :
				var NumControls = varDash.Controls.Count;
					
				if (varControl.Checked) {
					strSelectedItemValues = varControl.Text;
				}
				break;
				
			case bqTextBox :
				strSelectedItemValues = varControl.Text;
				break;
				
			case bqCheckBox :
				if(varControl.Checked) {
					strSelectedItemValues = varControl.Text;
				}
				break;
				
			default :		//	Other control type.
				Console.Writeln(sFctn + strControlName + " is the name of a control that is of a type that is not evaluated.");
				strSelectedItemValues = sFctn + strControlName + " is the name of a control that is of a type that is not evaluated.";
				break;
		}
	}
	catch (e) {
		//	strControlName is not the name of a control.  Maybe it's the name of a group.
		var NumControls = varDash.Shapes.Count
		var blnGroupFound = false;
		for (i = 1; i <= NumControls; i++) {
			if (varDash.Shapes[i].Group == strControlName) {
				blnGroupFound = true;
				if (varDash.Shapes[i].Checked == true)
					strSelectedItemValues = varDash.Shapes[i].Name;
			}
		}
		if (!blnGroupFound) {
			Console.Writeln(sMsg);
			return(sMsg);
		}
	}

	return strSelectedItemValues;
};


ActiveDocument.wsdot_LoadListBox = function (strDash, strCtl, strResult, strSourceCol, strItem1, bSortDesc, strSortCol) {
/*****************************************************************************/
//		Deprecated:  	This function is being replaced by the Shape.SetList and Shape.SetValue methods.
/*****************************************************************************/
/*
	Function:	loadListBox()
	Type:		JavaScript 
	Purpose:	Populates list control (ListBox or DropDown) with items in Brio result set.
	Inputs:		strDash			The name of the dashboard section containing the list box control.
				strListBoxName	Name of the list control to load.
				strResult		Name of result or table to get values from.
				strSourceCol	Name of column in result or table to get values from.
				strItem1		Text to be displayed as first item in drop down or list box list
								For no default first item, pass an empty string: "" (default none)
				bSortDesc		Flag to reverse sort order to descending (default ascending)
				strSortCol		Name of column in result or table to use for sorting
	
	Returns:	string			message indicating success or what error occurred
	
	Modification History:
	Date		Developer	Description
	5/1/03		R Neill		Original
	2/28/08		D Pulse		Added the ability to populate a DropDown
	6/4/08		D Pulse		Added the ability to use a table
*/
	Application.LibraryLogger.info("wsdot_LoadListBox function is deprecated.");
	var sMsg, e;
	var sFctn = "wsdot_LoadListBox:  ";
	
	//	check for input errors
	//	strDash
	sMsg = sFctn + strDash + " is not the name of a Dashboard object.";
	try {
		var varDash = ActiveDocument.Sections[strDash];
	}
	catch (e) {
		//	dashboard doesn't exist
		Console.Writeln(sMsg);
		return(sMsg);
	}
	
	if (varDash.Type != bqDashboard) {
		//	not a dashboard
		Console.Writeln(sMsg);
		return(sMsg);
	}
	
	//	strCtl
	sMsg = sFctn + strCtl + " is not the name of a ListBox or DropDown control in " + strDash + ".";	
	try {
		var varControl = varDash.Shapes[strCtl];
	}
	catch (e) {
		//	listbox or dropdown doesn't exist
		Console.Writeln(sMsg);
		return(sMsg);
	}
	if (varControl.Type != bqListBox && varControl.Type != bqDropDown) {
		//	not a listbox or dropdown
		Console.Writeln(sMsg);
		return(sMsg);
	}
	
	//	strResult
	sMsg = sFctn + "This document does not contain a Result or Table named " + strResult + ".";
	try {
		var varResult = ActiveDocument.Sections[strResult];
	}
	catch (e) {
		//	result or table doesn't exist
		Console.Writeln(sMsg);
		return(sMsg);
	}
	
	if (varResult.Type != bqResults && varResult.Type != bqTable) {
		//	not a result or table
		Console.Writeln(sMsg);
		return(sMsg);
	}
	
	//	strSourceCol
	try {
		var colSource = varResult.Columns[strSourceCol];
	}
	catch (e) {
		//	column doesn't exist
		sMsg = sFctn + strResult + " does not contain a column named " + strSourceCol + ".";
		Console.Writeln(sMsg);
		return(sMsg);
	}
	
	//	strSortCol
	if (strSortCol != undefined) {
		try {
			var colSort = varResult.Columns[strSortCol];
		}
		catch (e) {
			//	column doesn't exist
			sMsg = sFctn + strResult + " does not contain a column named " + strSortCol + ".";
			Console.Writeln(sMsg);
			return(sMsg);
		}
	}
	
	//	input is good -- continue
	
	var intItemCount = varResult.RowCount;
	var varTemp;

	varControl.RemoveAll();
	
	// Create blank item at top of list, if bFirstItemBlank.
	if (strItem1 != undefined && strItem1.length != 0) {
		Application.LibraryLogger.info("wsdot_LoadListBox:  Adding arbitrary first item.");
		varDash.Shapes[strCtl].Add(strItem1);
	}
	
	// Populate list control.
	if (strSortCol != undefined) {
		Application.LibraryLogger.info("wsdot_LoadListBox:  Sorting by alternate column.");
		
		with (varResult.SortItems) {
			RemoveAll();
			Add(strSortCol);
			Item(1).SortOrder = (bSortDesc ? bqSortDescend : bqSortAscend);
			SortNow();
		}
		colSource = varResult.Columns.Item(strSourceCol);
		for (j = 1; j <= varResult.RowCount; j++) {
			varControl.Add(colSource.GetCell(j));
		}
	}
	else {
		var tmpLim = varResult.Limits.CreateLimit(strSourceCol);
		tmpLim.RefreshAvailableValues();
		
		if (bSortDesc) {
			Application.LibraryLogger.info("wsdot_LoadListBox:  Sorting in descending order.");
			for (j = tmpLim.AvailableValues.Count; j >= 1; j--) {
				varControl.Add(tmpLim.AvailableValues[j]);
			}
		}
		else {
			for (j = 1; j <= tmpLim.AvailableValues.Count; j++) {
				varControl.Add(tmpLim.AvailableValues[j]);
			}
		}
	}

	// Select first item, if bFirstItemBlank.
	if (strItem1 != undefined && strItem1.length > 0) {
		varControl.Select(1);
	}
	return("Success");
};


ActiveDocument.wsdot_LoadListBox2 = function (oShape, oSourceCol, sItem1, iSortOrder, oSortCol) {
/*****************************************************************************/
//		Deprecated:  	This function is being replaced by the Shape.SetList and Shape.SetValue methods.
/*****************************************************************************/
/*
	Function:	loadListBox()
	Type:		JavaScript 
	Purpose:	Populates list control (ListBox or DropDown) with items in Brio result set.
	Inputs:		oShape			List control to load.
				oSourceCol		Column in result or table to get values from.
				sItem1			Text to be displayed as first item in drop down or list box list
								For no default first item, pass an empty string: "" (default none)
				iSortOrder		bqSortOrder value		1 = ascending, 2 = descending
				oSortCol		Column in result or table to use for sorting
	
	Returns:	string			message indicating success or what error occurred
	
	Modification History:
	Date		Developer	Description
	5/1/03		R Neill		Original
	2/28/08		D Pulse		Added the ability to populate a DropDown
	6/4/08		D Pulse		Added the ability to use a table
*/
	Application.LibraryLogger.info("wsdot_LoadListBox2 function is deprecated.");
	var arr = new Array();
	
	// Create blank item at top of list, if bFirstItemBlank.
	if (sItem1 != undefined && sItem1.length != 0) arr.push(sItem1);
	
	// Populate list control.
	if (oSortCol != undefined) {
		with (oSortCol.Parent.Parent.SortItems) {
			RemoveAll();
			Add(oSortCol.Name);
			Item(1).SortOrder = iSortOrder;
			SortNow();
		}
		for (j = 1; j <= oSourceCol.Parent.Parent.RowCount; j++) arr.push(oSourceCol.GetCell(j));
	}
	else {
		arr.concat(oSourceCol.Parent.Parent.getValues(oSourceCol.Name));
		if (iSortOrder == 2) arr.reverse();
	}
	oShape.SetList(arr);

	// Select first item
	if (sItem1 != undefined && sItem1.length > 0) oShape.Select(1);
	
	return("Success");
};


ActiveDocument.wsdot_AddSideLabels = function (strSection, strItem, bTotalRow, bInitializeRowTotals) {
/*
	Function:	fnAddSideLabels()
	
	Type:		JavaScript 
	Purpose:	Adds items to serve as side labels 
	Inputs:		strItem					The item in the given topic to use in limit.
				strSection				Name of query section to get topic, item from.
				bTotalRow				Add row total, if true
				bInitializeRowTotals	Clear existing row totals, if true.
	
	Returns:	None
	Modification History:
	Date		Developer	Description
	5/30/2003	Rich Neill	Original
*/
	var bSideLabelExists = true;
	
	try {
		var intSectionType = ActiveDocument.Sections[strSection].Type;
	}
	catch (e) {
		Console.Writeln(strSection + " is not the name of a Pivot object.");
		return(strSection + " is not the name of a Pivot object.");
	}
	if (ActiveDocument.Sections[strSection].Type != bqPivot) {
		Console.Writeln(strSection + " is not the name of a Pivot object.");
		return(strSection + " is not the name of a Pivot object.");
	}
	
	try {
		var strSideLabel = ActiveDocument.Sections[strSection].SideLabels.Item(strItem).Name;
	}
	catch (e) {
		//        item was not found -- good
		bSideLabelExists = false;
	}
	if (bSideLabelExists) {
		Console.Writeln(strItem + " is already a side label of " + strSection + ".");
		return(strItem + " is already a side label of " + strSection + ".");
	}
	
	try {
		ActiveDocument.Sections[strSection].SideLabels.Add(strItem);
	}
	catch (e) {
		Console.Writeln(strItem + " is not the name of an Item in " + strSection + ".");
		return(strItem + " is not the name of an Item in " + strSection + ".");
	}
	
	if(bInitializeRowTotals || bInitializeRowTotals == undefined)
	{
		ActiveDocument.Sections[strSection].SideLabels[strItem].Totals.RemoveAll();
	}
	if(bTotalRow)
	{
		ActiveDocument.Sections[strSection].SideLabels[strItem].Totals.Add();
	}
	return("Success");
};


ActiveDocument.wsdot_AddTopLabels = function (strSection, strItem, bTotalColumn) {
/*
	Function:	fnAddTopLabels()
	
	Type:       JavaScript 
	Purpose:	Adds items to serve as top labels 
	Inputs:		strItem			The item in the given topic to use in limit.
				strSection		Name of query section to get topic, item from.
				bTotalColumn	Add column total, if true
	
	Returns:	None
	Modification History:
	Date		Developer	Description
	5/30/2003	Rich Neill	Original
*/
	var bTopLabelExists = true;
	
	try {
		var intSectionType = ActiveDocument.Sections[strSection].Type;
	}
	catch (e) {
		Console.Writeln(strSection + " is not the name of a Pivot object.");
		return(strSection + " is not the name of a Pivot object.");
	}
	if (ActiveDocument.Sections[strSection].Type != bqPivot) {
		Console.Writeln(strSection + " is not the name of a Pivot object.");
		return(strSection + " is not the name of a Pivot object.");
	}
	
	try {
		var strTopLabel = ActiveDocument.Sections[strSection].TopLabels.Item(strItem).Name;
	}
	catch (e) {
		//        item was not found -- good
		bTopLabelExists = false;
	}
	if (bTopLabelExists) {
		Console.Writeln(strItem + " is already a top label of " + strSection + ".");
		return(strItem + " is already a top label of " + strSection + ".");
	}
	
	try {
		ActiveDocument.Sections[strSection].TopLabels.Add(strItem);
	}
	catch (e) {
		Console.Writeln(strItem + " is not the name of an Item in " + strSection + ".");
		return(strItem + " is not the name of an Item in " + strSection + ".");
	}
	
	if(bInitializeRowTotals || bInitializeRowTotals == undefined)
	{
		ActiveDocument.Sections[strSection].TopLabels[strItem].Totals.RemoveAll();
	}
	if(bTotalRow)
	{
		ActiveDocument.Sections[strSection].TopLabels[strItem].Totals.Add();
	}
	return("Success");
};


ActiveDocument.wsdot_SetLimitFromListbox = function (strEISSection, strControlName, strSection, strItem, strTopic) {
/*
	Function:	fnSetLimitFromListbox()
	
	Type:		JavaScript 
	Purpose:	Sets query limits based on user selection criteria in listboxes.
	Inputs:		strEISSection	The dashboard section to get the control values from.
				strControlName	Name of the listbox control to get values from.
				strSection		Name of query section to get topic, item from.
				strTopic		(optional) The topic in the query to use in limit
				strItem			The item in the given topic to use in limit.
	
	Returns:	None
	Modification History:
	Date		Developer	Description
	5/30/2003	Rich Neill	Original
	2/28/2008	Doug Pulse	Made this function generic.  It now handles a listbox or dropdown
							defines a limit in a query or resultset.  Reordered arguments for
							consistency with other scripts
*/
	var sMsg, e;
	var varLimit, varDash, varCtl, varSctn, strColName, varItem, varTopic;
	
	//	validate input
	//	Is strEISSection a dashboard?
	sMsg = strEISSection + " is not the name of a Dashboard object.";
	try {
		varDash = ActiveDocument.Sections[strEISSection];
	}
	catch (e) {
		//	not a section
		Console.Writeln(sMsg);
		return(sMsg);
	}
	if (varDash.Type != bqDashboard) {
		//	not a dashboard
		Console.Writeln(sMsg);
		return(sMsg);
	}
	
	//	Is strControlName a listbox or dropdown?
	sMsg = strControlName + " is not the name of a ListBox or DropDown control in " + strEISSection + ".";
	try {
		varCtl = varDash.Shapes[strControlName];
	}
	catch (e){
		//	not a control in strEISSection
		Console.Writeln(sMsg);
		return(sMsg);
	}
	if (varCtl.Type != bqListBox && varCtl.Type != bqDropDown) {
		//	not a listbox or dropdown
		Console.Writeln(sMsg);
		return(sMsg);
	}
	
	//	Is strSection a query or results?
	sMsg = strSection + " is not the name of a Query, Results, or Table section in the ActiveDocument.";
	try {
		varSctn = ActiveDocument.Sections[strSection];
	}
	catch (e){
		//	not a section
		Console.Writeln(sMsg);
		return(sMsg);
	}
	if (varSctn.Type != bqQuery && varSctn.Type != bqResults && varSctn.Type != bqTable) {
		//	not a query, results, or table
		Console.Writeln(sMsg);
		return(sMsg);
	}
	
	if (varSctn.Type == bqQuery) {
		//	strSection is a query
		//	Is strTopic in strSection?
		try {
			varTopic = varSctn.DataModel.Topics[strTopic];
		}
		catch (e){
			//	not a topic
			sMsg = strTopic + " is not the name of a Topic in " + strSection + ".";
			Console.Writeln(sMsg);
			return(sMsg);
		}
		//	Is strItem in strTopic?
		try {
			varItem = varTopic.TopicItems[strItem];
		}
		catch (e){
			//	not an item
			sMsg = strItem + " is not the name of a column in " + strSection + ".";
			Console.Writeln(sMsg);
			return(sMsg);
		}
	}
	else {
		//	strSection is a results or table
		//	Is strItem in strSection?
		try {
			varItem = varSctn.Columns[strItem];
		}
		catch (e){
			//	not an item
			sMsg = strItem + " is not the name of a column in " + strSection + ".";
			Console.Writeln(sMsg);
			return(sMsg);
		}
	}
	
	
	//	input is good -- continue
	strColName = ((varSctn.Type == bqQuery) ? strTopic + "." : "") + strItem;
	
	varLimit = varSctn.Limits.CreateLimit(strColName);
	varLimit.Name = strItem;
	varLimit.DisplayName = strItem;
	varLimit.Operator = bqLimitOperatorEqual;
	
	if (varCtl.Type == bqListBox) {	//	we're using a listbox
		for(i = 1; i <= varCtl.SelectedList.Count; i++)
		{
			varLimit.CustomValues.Add(varCtl.SelectedList[i]);
			varLimit.SelectedValues.Add(varCtl.SelectedList[i]);	
		}
	}
	else {		//	we're"	"using a dropdown
		i = varCtl.SelectedIndex;
		varLimit.CustomValues.Add(varCtl.Item(i));
		varLimit.SelectedValues.Add(varCtl.Item(i));
	}
	varSctn.Limits.Add(varLimit);
	return("Success");	//	success
};


function GlobalStartup() {
//	Provide an object to...
	
	//	Allow the scripter to see if this class has been updated
	this.Updated = new Date("09/23/2013");
	
	//	Provide the scripter with an updated status messages
	this.getMsg = function () {
		var sTemp = "Three years since last update.  Lots of changes.  Contact lead developer for details.";
		sTemp += "";
		return sTemp;
	};
	
	//	Link to help
	this.getHelp = function () {
		var sApp = "C:\\Program Files\\Internet Explorer\\iexplore.exe";
		var sArgs = "http://datamining/Training/Hyperion/FunctionDoc/ScriptingDocumentation.htm";
		Application.Shell(sApp, sArgs);
	};
	
	//	List all scripts in the active document
	this.OutputScripts = ActiveDocument.OutputScripts;
	this.showScripts = this.OutputScripts;
	
	//	load scripts from a text file into the ActiveDocument
	this.InputScripts = ActiveDocument.InputScripts;
	
	//	In the WSDOT Dashboard, position toolbar buttons and update their comments (tooltips).
	this.setToolbarButtons = function () {
		for (var i = 1; i <= ActiveDocument.Sections.Count; i++) {
			if (ActiveDocument.Sections.Item(i).Type == bqDashboard) {
				for (var j = 1; j <= ActiveDocument.Sections.Item(i).Shapes.Count; j++) {
					if (ActiveDocument.Sections.Item(i).Shapes.Item(j).Type == bqPicture) {
						//if ((new Array("picHelp", "picHelpD", "picPrint", "picPrintD", "picPrintPreview", "picPrintPreviewD", "picSend", "picSendD", "picEdit", "picEditD", "picExport", "picExportD")).contains(ActiveDocument.Sections.Item(i).Shapes.Item(j).Name)) {
						//	ActiveDocument.Sections.Item(i).Shapes.Item(j).Placement.YOffset = 0;
						//	ActiveDocument.Sections.Item(i).Shapes.Item(j).Placement.Height = 26;
						//	ActiveDocument.Sections.Item(i).Shapes.Item(j).Placement.Width = 26;
							switch (ActiveDocument.Sections.Item(i).Shapes.Item(j).Name) {
								case "picHelp" :
									ActiveDocument.Sections.Item(i).Shapes.Item(j).Comments = "Help";
									ActiveDocument.Sections.Item(i).Shapes.Item(j).Placement.XOffset = 984 - 26;
									break;
								case "picHelpD" :
									ActiveDocument.Sections.Item(i).Shapes.Item(j).Comments = "Help";
									ActiveDocument.Sections.Item(i).Shapes.Item(j).Placement.XOffset = 984 - 26;
									break;
								case "picPrint" :
									ActiveDocument.Sections.Item(i).Shapes.Item(j).Comments = "Print";
									ActiveDocument.Sections.Item(i).Shapes.Item(j).Placement.XOffset = 984 - (2 * 26);
									break;
								case "picPrintD" :
									ActiveDocument.Sections.Item(i).Shapes.Item(j).Comments = "Print";
									ActiveDocument.Sections.Item(i).Shapes.Item(j).Placement.XOffset = 984 - (2 * 26);
									break;
								case "picPrintPreview" :
									ActiveDocument.Sections.Item(i).Shapes.Item(j).Comments = "Print Preview";
									ActiveDocument.Sections.Item(i).Shapes.Item(j).Placement.XOffset = 984 - (3 * 26);
									break;
								case "picPrintPreviewD" :
									ActiveDocument.Sections.Item(i).Shapes.Item(j).Comments = "Print Preview";
									ActiveDocument.Sections.Item(i).Shapes.Item(j).Placement.XOffset = 984 - (3 * 26);
									break;
								case "picSend" :
									ActiveDocument.Sections.Item(i).Shapes.Item(j).Comments = "Send";
									ActiveDocument.Sections.Item(i).Shapes.Item(j).Placement.XOffset = 984 - (4 * 26);
									break;
								case "picSendD" :
									ActiveDocument.Sections.Item(i).Shapes.Item(j).Comments = "Send";
									ActiveDocument.Sections.Item(i).Shapes.Item(j).Placement.XOffset = 984 - (4 * 26);
									break;
								case "picExport" :
									ActiveDocument.Sections.Item(i).Shapes.Item(j).Comments = "Export";
									ActiveDocument.Sections.Item(i).Shapes.Item(j).Placement.XOffset = 984 - (5 * 26);
									break;
								case "picExportD" :
									ActiveDocument.Sections.Item(i).Shapes.Item(j).Comments = "Export";
									ActiveDocument.Sections.Item(i).Shapes.Item(j).Placement.XOffset = 984 - (5 * 26);
									break;
								case "picEdit" :
									ActiveDocument.Sections.Item(i).Shapes.Item(j).Comments = "Edit";
									ActiveDocument.Sections.Item(i).Shapes.Item(j).Placement.XOffset = 984 - (6 * 26);
									break;
								case "picEditD" :
									ActiveDocument.Sections.Item(i).Shapes.Item(j).Comments = "Edit";
									ActiveDocument.Sections.Item(i).Shapes.Item(j).Placement.XOffset = 984 - (6 * 26);
									break;
								case "picAddFav" :
									ActiveDocument.Sections.Item(i).Shapes.Item(j).Comments = "Add to Favorites";
									ActiveDocument.Sections.Item(i).Shapes.Item(j).Placement.XOffset = 984 - (7 * 26);
									break;
								case "picAddFavD" :
									ActiveDocument.Sections.Item(i).Shapes.Item(j).Comments = "Add to Favorites";
									ActiveDocument.Sections.Item(i).Shapes.Item(j).Placement.XOffset = 984 - (7 * 26);
									break;
							}
						//}
					}
				}
			}
		}
	}
	
	//	In the all dashboards in the ActiveDocument, create a GetValue, SetValue, GetList, and SetList for each shape (if appropriate).
	this.loadObjectMethods = ActiveDocument.loadObjectMethods;
	
	this.SetToolTips = function () {
		var arrDash = new Array();
		var arrShape = new Array("picLogo", "picAddFav", "picExport", "picEdit", "picSend", "picPrint", "picPrintPreview", "picHelp");
		var arrShapeD = new Array("picLogoD", "picAddFavD", "picExportD", "picEditD", "picSendD", "picPrintD", "picPrintPreviewD", "picHelpD");
		var arrComment = new Array("WSDOT Intranet", "Add to Favorites", "Export", "Edit", "Send", "Print", "Print Preview", "Help");
		var dctComment = new wsdot_Dictionary();
		for (var i = 0; i < arrShape.length; i++) {
			dctComment.Add(arrShape[i], arrComment[i]);
		}
		
		for (var i = 1; i < ActiveDocument.Sections.Count; i++) {
			if (ActiveDocument.Sections.Item(i).Type == bqDashboard) {
				arrDash.push(ActiveDocument.Sections.Item(i));
			}
		}
		
		for (var i = 0; i < arrDash.length; i++) {
			for (var j = 1; j < arrDash[i].Shapes.Count; j++) {
				if (arrShape.contains(arrDash[i].Shapes.Item(j).Name)) {
					arrDash[i].Shapes.Item(j).Comments = dctComment.Item(arrDash[i].Shapes.Item(j).Name);
				}
				if (arrShapeD.contains(arrDash[i].Shapes.Item(j).Name)) {
					arrDash[i].Shapes.Item(j).Comments = "";
				}
			}
		}
	};
};
ActiveDocument.wsdot_GlobalStartup = new GlobalStartup();

ActiveDocument.g_blnHelp = false;
ActiveDocument.g_blnDebug = false;

ActiveDocument.g_fsoMain = Application.FSO;
ActiveDocument.g_strUserTempPath = Application.TempPath;
ActiveDocument.g_strDocPath = Application.DocPath;
ActiveDocument.g_strWinPath = Application.WinPath;
ActiveDocument.g_strUserName = Application.UserName;
ActiveDocument.g_strComputerName = Application.ComputerName;

ActiveDocument.WeekdayName = Application.WeekdayName;
ActiveDocument.WeekdayAbbrev = Application.WeekdayAbbrev;
ActiveDocument.MonthName = Application.MonthName;
ActiveDocument.MonthAbbrev = Application.MonthAbbrev;


var blnTest = (ActiveDocument.ReposPath.toLowerCase().indexOf("test") != -1)
for (var i = 1; i <= ActiveDocument.Sections.Count; i++) {
	if (ActiveDocument.Sections.Item(i).Type == bqDashboard) {
		if (blnTest) {
			try {
				ActiveDocument.Sections.Item(i).Shapes["lblTest"].Visible = true;
			}
			catch (e) { }
		}
		else {
			try {
				ActiveDocument.Sections.Item(i).Shapes["lblTest"].Visible = false;
			}
			catch (e) { }
		}
	}
}

eval( (new JOOLEObject("Scripting.FileSystemObject")).OpenTextFile(-2, 1, "\\\\hqolymhyperion2\\SharedResources\\Script\\msg.js").ReadAll() );
//eval( (new JOOLEObject("Scripting.FileSystemObject")).OpenTextFile(-2, 1, "C:\\DMSVN\\SharedResources\\Reporting Tool\\Trunk\\msg.js").ReadAll() );

for (var i = 0; i < RepositoryMessages.length; i++) {
	var sPath = RepositoryMessages[i][0];
	var sHead = RepositoryMessages[i][1];
	var sMsg = RepositoryMessages[i][2];
	if ((new Array("/", "\\")).contains(sPath.charAt(sPath.length - 1))) {
		//	this is a folder
		if (ActiveDocument.Folder.indexOf(sPath) > -1 && sHead.length > 0) Alert(sMsg, sHead);
	}
	else {
		//	this is a file
		if (sPath == ActiveDocument.ReposPath && sHead.length > 0) Alert(sMsg, sHead);
	}
}

DoEvents();

