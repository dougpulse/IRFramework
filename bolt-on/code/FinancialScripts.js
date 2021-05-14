//	File:		FinancialScripts.js
//	Language:	JavaScript
//	Purpose:	to provide common scripts that apply to the user interfaces that have been 
//				developed for the Financial Data Mart but are not common to all files using scripting
//	Author:		doug pulse
//	Revision History:
//	doug pulse		5/20/2010	inaugural
//

ActiveDocument.RunQueries = function (arrQuery, arrListName, arrLimitName, dash, o, blnFiscal, pb, blnFilterBien, arrProcess) {
	//	Process:	RunQuery()
	//	Language:	JavaScript
	//	Purpose:	Set limits and run each query
	//	Inputs:		arrQuery		Array containing the Query sections to process
	//				arrListName		Array containing the names of the objects containing the limit criteria
	//				arrLimitName	Array containing the names of the limits
	//				o				the object calling this function
	//				dash			parent dashboard of the object calling this function
	//				blnFiscal		boolean value indicating whether fiscal periods are being used
	//								false = calendar periods
	//				pb				(optional) progressbar to use
	//				blnFilterBien	
	//	

	//ProgressBar = function (dashboard, left, top, width, height, value, maxvalue)
	//dash.GoProgress = new wsdot_ProgressBar(dash, dash.lblCriteria.Placement.XOffset, dash.lblCriteria.Placement.YOffset + o.Placement.Height, dash.lblCriteria.Placement.Width, o.Placement.Height, 0, 100);
	//dash.GoProgress = new wsdot_ProgressBar(dash, dash.lblCriteria.Placement.XOffset, o.Placement.YOffset + o.Placement.Height + 8, dash.lblCriteria.Placement.Width, o.Placement.Height, 0, 100);
	if (typeof pb != "undefined") {
		pb.Show();
	}
	
	if (typeof blnFilterBien != "boolean") {
		blnFilterBien = false;
	}
	
	if (blnFilterBien) {
		var arrBienYr = ActiveDocument.Sections["Selection Criteria"].BeginBienYr();
	}
	DoEvents();
	
	//Console.Writeln(blnFilterBien.toString());
	//Console.Writeln(arrQuery.length);
	
	//	determine which type of fiscal period is being requested (fiscal or calendar)
	sOn = (blnFiscal ? "Fiscal " : "Calendar ");
	sOff = (blnFiscal ? "Calendar " : "Fiscal ");
	
	for (var k = 0; k < arrQuery.length; k++) {
		//Console.Writeln(arrQuery[k].Name);
		dash.lblCriteria.Text = "Processing:  " + arrQuery[k].Name.substr(2);
		//	ignore the fiscal periods by default
		try {
			arrQuery[k].Limits["Fiscal Year"].Ignore = true;
			arrQuery[k].Limits["Fiscal Month Name"].Ignore = true;
			arrQuery[k].Limits["Calendar Year"].Ignore = true;
			arrQuery[k].Limits["Calendar Month Name"].Ignore = true;
		}
		catch (e) {
			//Console.Writeln("error 1");
		}
		
		//Console.Writeln(dash.arrListName.length);
		for (var i = 0; i < arrListName.length; i++) {
			//Console.Write(arrListName[i].column(16));
			try {
				var strLimName = arrLimitName[i];
				if ((new Array("Year", "MoName")).contains(arrListName[i])) {
					strLimName = sOn + (arrListName[i] == "Year" ? "Year" : "Month Name");
				}
				//Console.Write(strLimName.column(24));
				try { arrQuery[k].Limits[strLimName].Ignore = true; }
				catch (e) { }
				
				//	Is the operator "null" or "not null"?
				if ((new Array(12, 13)).contains(ActiveDocument.Sections["Selection Criteria"].Shapes["drpOper" + arrListName[i]].SelectedIndex)) {
					//Console.Writeln("null".column(16));
					//	yes
					//		don't ignore
					try {arrQuery[k].Limits[strLimName].Ignore = false;}
					catch (e) {
						//Console.Writeln("error 2");
					}
					
					try {
						//		set the operator to null (12)
						arrQuery[k].Limits[strLimName].Operator = 12;
						//		set the negate (true if "not null")
						arrQuery[k].Limits[strLimName].Negate = ActiveDocument.Sections["Selection Criteria"].Shapes["drpOper" + arrListName[i]].SelectedIndex == 13;
					}
					catch (e) {
						//Console.Writeln("error 3");
					}
				}
				else {
					//Console.Write("not null".column(16));
					//	no
					//		Is the criteria blank ("")?
					if (ActiveDocument.Sections["Selection Criteria"].Shapes["txt" + arrListName[i]].Text == "") {
						//	yes
						//Console.Writeln("blank".column(20));
						//		ignore
							//	this is the default
					}
					else {
						//	no
						//Console.Writeln(ActiveDocument.Sections["Selection Criteria"].Shapes["txt" + arrListName[i]].Text.column(20));
						//		don't ignore
						try {
							var arr = new Array();
							var arr2 = new Array();
							var arr3 = new Array();
							
							arrQuery[k].Limits[strLimName].Ignore = false;
							//		set the operator to the value in the dropdown
							arrQuery[k].Limits[strLimName].Operator = ActiveDocument.Sections["Selection Criteria"].Shapes["drpOper" + arrListName[i]].SelectedIndex;
							//		set negate to false
							arrQuery[k].Limits[strLimName].Negate = false;
							//		clear the custom list
							arrQuery[k].Limits[strLimName].CustomValues.RemoveAll();
							//		clear the selected list
							arrQuery[k].Limits[strLimName].SelectedValues.RemoveAll();
							//		add the criteria values to the custom list
							//		add the criteria values to the selected list
							var sFilter = ActiveDocument.Sections["Selection Criteria"].Shapes["txt" + arrListName[i]].Text;
							arr2 = sFilter.split("\"");
							
							if (arr2.length != 1) {
								for (var j = 0; j < arr2.length; j++) {
									if (j % 2 == 0) {
										//	not in quotes
										arr3 = arr2[j].split(",");
										for (var m = 0; m < arr3.length; m++) {
											if (arr3[m] != "") {
												arr.push(arr3[m]);
											}
										}
									}
									else {
										//	in quotes
										if (arr2[j] != "") {
											arr.push(arr2[j]);
										}
									}
								}
							}
							else {
								arr = sFilter.split(",");
							}
							
							for (var j = 0; j < arr.length; j++) {
								arrQuery[k].Limits[strLimName].CustomValues.Add(arr[j]);
								arrQuery[k].Limits[strLimName].SelectedValues.Add(arr[j]);
							}
						}
						catch (e) {
							//Console.Writeln("error 4");
						}
					}
				}
			}
			catch (e) {
				Console.Writeln("    Error:  " + e);
			}
		}
		//Console.Writeln("done");
		//Console.Writeln("");
		DoEvents();
		
		if (blnFilterBien) {
		//Console.Writeln("arrBienYr:  " + arrBienYr.length);
			//	if there is no begin biennium year filter yet, set it
			if (arrBienYr[0] == 0) {
				//	do nothing
			}
			else {
				var oLimBien = arrQuery[k].Limits["Fiscalbienniumid"];
				oLimBien.Ignore = false;
				oLimBien.CustomValues.RemoveAll();
				oLimBien.SelectedValues.RemoveAll();
				for (var i = 0; i < arrBienYr.length; i++) {
					oLimBien.CustomValues.Add(arrBienYr[i]);
					oLimBien.SelectedValues.Add(arrBienYr[i]);
				}
			}
		}
		
		if (typeof arrProcess == "undefined") {
			arrQuery[k].Process();
		}
		else if (arrProcess[k]) {
			arrQuery[k].Process();
		}
		
		if (typeof pb != "undefined") {
			pb.Value((100 / arrQuery.length) * (k + 1));
		}
		DoEvents();
	}
	if (typeof pb != "undefined") {
		pb.Kill();
	}
	//Console.Writeln("______________________________________________");
	//Console.Writeln("");
}



ActiveDocument.SetPart = function () {
	//	Purpose:	set the value of the partition filter
	var n = 0;
	var blnBien = true;
	var blnYr = true;
	var blnMo = true;
	
	var t = new Array();		//	textboxes holding filter values
	var o = new Array();		//	operator dropdowns
	try {
		t.push(ActiveDocument.Sections["Selection Criteria"].Shapes["txtBien"]);
		o.push(ActiveDocument.Sections["Selection Criteria"].Shapes["drpOperBien"]);
	}
	catch (e) {
		//	this file doesn't use a biennium filter
		blnBien = false;
	}
	try {
		t.push(ActiveDocument.Sections["Selection Criteria"].Shapes["txtYear"]);
		o.push(ActiveDocument.Sections["Selection Criteria"].Shapes["drpOperYear"]);
	}
	catch (e) {
		//	this file doesn't use a year filter
		blnYr = false;
	}
	try {
		t.push(ActiveDocument.Sections["Selection Criteria"].Shapes["txtMoName"]);
		o.push(ActiveDocument.Sections["Selection Criteria"].Shapes["drpOperMoName"]);
	}
	catch (e) {
		//	this file doesn't use a month name filter
		blnMo = false;
	}
	
	var q = new Array();		//	limits to use to take advantage of partitioning
	for (var i = 0; i < ActiveDocument.arrQuery.length; i++) {
		if (ActiveDocument.arrQueryPart[i]) q.push(ActiveDocument.arrQuery[i].Limits["Fiscalbienniumid"]);
	}
	//q.push(ActiveDocument.Sections["Q-Work Order Authorized Amounts"].Limits["Fiscalbienniumid"]);
	//q.push(ActiveDocument.Sections["Q-Expenditure Summary"].Limits["Fiscalbienniumid"]);
	//q.push(ActiveDocument.Sections["Q-Revenue by Work Order"].Limits["Fiscalbienniumid"]);
	//q.push(ActiveDocument.Sections["Q-Expenditure Detail"].Limits["Fiscalbienniumid"]);
	
	var d = ActiveDocument.Sections["r-Fiscal Month"];
	
	var aF = new Array();		//	names of fiscal period limits in [r-Fiscal Month]
	if (blnBien) aF.push("Biennium");
	if (blnYr) aF.push("Fiscal Year");
	if (blnMo) aF.push("Fiscal Month Name");
	
	var aC = new Array();		//	names of calendar period limits in [r-Fiscal Month]
	if (blnBien) aC.push("Biennium");
	if (blnYr) aC.push("Calendar Year");
	if (blnMo) aC.push("Calendar Month Name");
	
	var a = new Array();		//	names of limits to use in [r-Fiscal Month]
	try {
		a = ActiveDocument.Sections["Selection Criteria"].Shapes["rdoFiscal"].Checked ? aF : aC;
	}
	catch (e) {
		//	this file doesn't allow a fiscal/calendar selection
		a = aF;
	}
	
	var l = new Array();		//	limit objects to use in [r-Fiscal Month]
	for (var i = 0; i < a.length; i++) {
		try { l[i] = d.Limits[a[i]]; }
		catch (e) { }
	}
	
	//	adjust filters to take advantage of partitioning
	//	don't ignore, and remove all filter values from the partition filter
	for (var i = 0; i < q.length; i++) {
		q[i].Ignore = false;
		q[i].CustomValues.RemoveAll();
		q[i].SelectedValues.RemoveAll();
	}
	
	for (var i = 0, s = new Array(); i < t.length; i++) {
		if (t[i].Text.length != 0) {
			s = t[i].Text.split(",");
			l[i].Ignore = false;
			l[i].Operator = o[i].SelectedIndex;
			l[i].CustomValues.RemoveAll();
			l[i].SelectedValues.RemoveAll();
			for (var j = 0; j < s.length; j++) {
				l[i].CustomValues.Add(s[j]);
				l[i].SelectedValues.Add(s[j]);
			}
			n++;
		}
	}
	
	if (n == 0) {
		for (var j = 0; j < q.length; j++) {
			q[j].Ignore = true;
		}
	}
	else {
		for (var j = 0; j < q.length; j++) {
			if (d.RowCount == 0) {
				q[j].CustomValues.Add(-1);
				q[j].SelectedValues.Add(-1);
			}
			else {
				for (var k = 1; k <= d.RowCount; k++) {
					q[j].CustomValues.Add(d.Columns["Begin Biennium Year"].GetCell(k));
					q[j].SelectedValues.Add(d.Columns["Begin Biennium Year"].GetCell(k));
				}
			}
		}
	}
	
	for (var i = 0; i < l.length; i++) l[i].Ignore = true;
}


//	could be replaced by 
//ActiveDocument.ExportFinancialData = function (oDash, sResultPrefix) {
//	ActiveDocument.ExportData(oDash, sResultPrefix, "Financial Reports");
//}
//ActiveDocument.ExportData = function (oDash, sResultPrefix) {
//	//	Purpose:  export the data from the report to pdf or from the results to excel
//	var d = new Date();
//	var dfile = new String("");
//	var iOutFormat, sSection;
//	var sBaseName = oDash.Name;
//	
//	if (oDash.ActiveTab) {
//		//	to accommodate multiple levels of tabs
//		sBaseName += " - " + oDash.ActiveTab;
//	}
//	
//	if (oDash.FPType) {
//		//	to allow file naming based on the type of fiscal period
//		sBaseName = sBaseName.replace(/ Year/gi, " " + oDash.FPType + " Year");
//	}
//	
//	dfile = ActiveDocument.g_strUserTempPath + "\\" + sBaseName + "_" + d.format("yyyymmddhhnnss");
//	
//	var iOut = Alert("What do you want to export?", "Financial Reports", "Printable Document", "Data to Spreadsheet", "Cancel");
//	switch (iOut) {
//		case 1 :
//			dfile += ".pdf";
//			iOutFormat = bqExportFormatPDF;
//			if (oDash.ActiveTab) {
//				sSection = sBaseName;
//			}
//			else {
//				sSection = sBaseName.replace(/ /gi, "  ");
//			}
//			ActiveDocument.Sections[sSection].Export(dfile, iOutFormat, true, false);
//			Shell(g_strWinPath + "\\explorer.exe", "\"" + dfile + "\"");
//			break;
//			
//		case 2 :
//			sSection = sResultPrefix + "-" + sBaseName;
//			var iOut2 = 2;
//			
//			if (ActiveDocument.Sections[sSection].Type == bqPivot) {
//				iOut2 = Alert("Do you want the cells unmerged?", "Financial Reports", "Yes", "No", "Cancel");
//			}
//			
//			dfile = sBaseName + "_" + d.format("yyyymmddhhnnss");
//			
//			switch (iOut2) {
//				case 1 :
//					ActiveDocument.Sections[sSection].ExportToExcel(ActiveDocument.g_strUserTempPath, dfile, true, true);
//					break;
//					
//				case 2 :
//					ActiveDocument.Sections[sSection].ExportToExcel(ActiveDocument.g_strUserTempPath, dfile, true, false);
//					break;
//					
//				default :
//					break;
//			}
//			
//			break;
//			
//		case 3 : 
//			//	user cancelled
//			break;
//			
//		default :
//			break;
//	}
//}



ActiveDocument.DropDown = function (o) {
	//	Purpose:	manage what is shown or hidden when the user clicks on a dropdown image
	o.Parent.Shapes["rctBack"].OnClick();
	o.Parent.sCurrCtrl = o.Name.substr(3);
	for (var i = o.Parent.arrListName.indexOf(o.Parent.sCurrCtrl); i < o.Parent.arrListName.length; i++) {
		o.Parent.Shapes["txt" + o.Parent.arrListName[i]].Visible = false;
		o.Parent.Shapes["pic" + o.Parent.arrListName[i]].Visible = false;
	}
	o.Parent.Shapes["btnGo"].Visible = false;
	o.Parent.Shapes["lst" + o.Parent.sCurrCtrl].Visible = true;
	o.Parent.Shapes["btnSet"].Placement.YOffset = o.Placement.YOffset;
	o.Parent.Shapes["btnSet"].Visible = true;
}



//	clear results
//ActiveDocument.ClearResults(ActiveDocument.arrQuery, ActiveDocument.arrProcess);

//if (!ActiveDocument.blnCleared) {
//	for (var i = 0; i < arrQuery.length; i++) {
//		if (typeof ActiveDocument.arrProcess == "undefined") {
//			ActiveDocument.arrQuery[i].Limits["Id"].Ignore = false;
//			ActiveDocument.arrQuery[i].Process();
//			ActiveDocument.arrQuery[i].Limits["Id"].Ignore = true;
//		}
//		else if (ActiveDocument.arrProcess[i]) {
//			ActiveDocument.arrQuery[i].Limits["Id"].Ignore = false;
//			ActiveDocument.arrQuery[i].Process();
//			ActiveDocument.arrQuery[i].Limits["Id"].Ignore = true;
//		}
//	}
//}



ActiveDocument.wsdot_GlobalStartup.loadObjectMethods();

//	added 2/27/2013 by doug pulse
ActiveDocument.HomeDashboard = "Selection Criteria";
ActiveDocument.PreloadHomeSection = true;
//

//ActiveDocument.Sections["Selection Criteria"].Shapes["btnSetCleanScript"].OnClick();

//	if the file is loaded in the web client, don't prompt to save when the document is closed
if ((Application.Type == bqAppTypePlugInClient) || (Application.Type == bqAppTypeThinClient))
	ActiveDocument.PromptToSave = false;


//	if the file is loaded in the desktop client, make sure all queries are connected to the database
//if (Application.Type == bqAppTypeDesktopClient) {
	var SecCount = ActiveDocument.Sections.Count;
	for (j = 1; j <= SecCount ; j++) {
		if (ActiveDocument.Sections[j].Type == bqQuery) {
			try {
				ActiveDocument.Sections[j].DataModel.Connection.Connect();
			}
			catch (e) {
				Console.Writeln(e.toString());
				Console.Writeln("    " + ActiveDocument.Sections[j].Name);
			}
		}
	}
//}


//	process the queries needed to take advantage of partitioning
ActiveDocument.Sections["q-Partitions"].Process();
ActiveDocument.Sections["q-Fiscal Month"].Process();


////	get partition information
//ActiveDocument.dctPartition = new wsdot_Dictionary();
//for (var i = 1, s = "", a = new Array(); i <= ActiveDocument.Sections["r-Partitions"].RowCount; i++) {
//	s = ActiveDocument.Sections["r-Partitions"].Columns["Databasename"].GetCell(i);
//	s += ".";
//	s += ActiveDocument.Sections["r-Partitions"].Columns["Facttablename"].GetCell(i);
//	a.push(ActiveDocument.Sections["r-Partitions"].Columns["Factcolumnname"].GetCell(i));
//	a.push(ActiveDocument.Sections["r-Partitions"].Columns["Dimensiontablename"].GetCell(i));
//	a.push(ActiveDocument.Sections["r-Partitions"].Columns["Dimensioncolumnname"].GetCell(i));
//	dctPartition.Add(s, a);
//}

//	create month arrays to use for monthname dropdown
ActiveDocument.arrCalMonths = ActiveDocument.MonthName.concat("Month99").concat("Month25");
ActiveDocument.arrFY1Months = ActiveDocument.MonthName.concat("Month99").sort(SortFY1MonthNames);
ActiveDocument.arrFY2Months = ActiveDocument.MonthName.concat("Month25").sort(SortFY2MonthNames);
ActiveDocument.arrFiscalMonths = ActiveDocument.arrFY1Months.concat("Month25");


//	set the default value for the biennium criteria [to the current biennium]
var a = new Array();
var blnBien = false;
try { blnBien = (ActiveDocument.Sections["Selection Criteria"].lstBien.Type == bqListBox) }
catch (e) {
	//	lstBien is not a ListBox -- do nothing
}


if (blnBien) {
	var oL = ActiveDocument.Sections["r-Fiscal Month"].Limits["Biennium"];
	oL.RefreshAvailableValues();
	var oAV = oL.AvailableValues;
	var _arr = new Array();
	//Console.Writeln(oAV.Count);
	
	//Console.Writeln(oL.Name);
	for (var i = 1; i <= oAV.Count; i++) {
		//Console.Writeln("    " + oAV.Item(i));
		if (oAV.Item(i).toString().substr(0, 2) != "DW") {
			_arr.push(oAV.Item(i));
			//Console.Writeln("    \'" + oAV.Item(i) + "\'");
		}
	}
	
	ActiveDocument.Sections["Selection Criteria"].lstBien.SetList(_arr.reverse());
	
	for (var i = 1; i <= ActiveDocument.Sections["Selection Criteria"].lstBien.SelectedList.Count; i++) {
		a.push(ActiveDocument.Sections["Selection Criteria"].lstBien.SelectedList.ItemIndex(i));
	}
	for (var i = 0; i <= a.length; i++) {
		ActiveDocument.Sections["Selection Criteria"].lstBien.Unselect(a[i]);
	}
	a = ActiveDocument.Sections["Selection Criteria"].lstBien.GetList();
	
	ActiveDocument.Sections["Selection Criteria"].lstBien.Select(a.indexOf((new Date()).getBiennium()) + 1);
	try {
		ActiveDocument.Sections["Selection Criteria"].txtBien.Text = (new Date()).getBiennium();
	}
	catch (e) {
		//	meaningless error
	}
}


//	set the default value for the year criteria [to the current fiscal year]
a.slice(0, 0);
var blnYear = false;
try { blnYear = (ActiveDocument.Sections["Selection Criteria"].lstYear.Type == bqListBox) }
catch (e) {
	//	lstYear is not a ListBox -- do nothing
}

if (blnYear) {
	var oL = ActiveDocument.Sections["r-Fiscal Month"].Limits["Fiscal Year"];
	oL.RefreshAvailableValues();
	var oAV = oL.AvailableValues;
	var _arr = new Array();
	
	//Console.Writeln(oL.Name);
	for (var i = 1; i <= oAV.Count; i++) {
		//Console.Writeln("    " + oAV.Item(i));
		if (oAV.Item(i).toString().substr(0, 2) != "DW") {
			_arr.push(oAV.Item(i));
			//Console.Writeln("    \'" + oAV.Item(i) + "\'");
		}
	}
	
	ActiveDocument.Sections["Selection Criteria"].lstYear.SetList(_arr.reverse());
	
	for (var i = 1; i <= ActiveDocument.Sections["Selection Criteria"].lstYear.SelectedList.Count; i++) {
		a.push(ActiveDocument.Sections["Selection Criteria"].lstYear.SelectedList.ItemIndex(i));
	}

	for (var i = 0; i <= a.length; i++) {
		ActiveDocument.Sections["Selection Criteria"].lstYear.Unselect(a[i]);
	}
	a = ActiveDocument.Sections["Selection Criteria"].lstYear.GetList();
	
	ActiveDocument.Sections["Selection Criteria"].lstYear.Select(a.indexOf((new Date()).getFiscalYear().toString()) + 1);
	//ActiveDocument.Sections["Selection Criteria"].lstYear.Select((new Date()).getFiscalYear() - ActiveDocument.Sections["Selection Criteria"].lstYear.Item(1) + 1);
	try {
		ActiveDocument.Sections["Selection Criteria"].txtYear.Text = (new Date()).getFiscalYear();
	}
	catch (e) {
		//	meaningless error
	}
}



var sScript = "var strInput = this.Text;\r\n";
sScript += "for (var i = 0; i < strInput.length; i++) {\r\n";
sScript += "	if ((strInput.charCodeAt(i) == 13) && (strInput.charCodeAt(i + 1) == 10)) {\r\n";
sScript += "		i++;\r\n";
sScript += "		var strOutput = this.Text.replace(/\\r\\n/g, \"\");\r\n";
sScript += "		this.Text = strOutput;\r\n";
sScript += "		DoEvents();\r\n";
sScript += "	}\r\n";
sScript += "}\r\n\r\n";

var o;		//	textbox shape
for (var j = 0; j < ActiveDocument.Sections["Selection Criteria"].arrListName.length; j++) {
	o = ActiveDocument.Sections["Selection Criteria"].Shapes["txt" + ActiveDocument.Sections["Selection Criteria"].arrListName[j]];
	sTemp = o.EventScripts["OnChange"].Script;
	if (sTemp.search(/^var strInput/gi) == -1) {
		o.EventScripts["OnChange"].Script = sScript + sTemp;
	}
}



//	put a disclaimer on each dashboard and report
//	remove this when we go live
//	3/6/2012 -- remove disclaimers
var d, lTop, lLeft, s;
var myFont = new Application.Font();

for (var i = 1; i <= ActiveDocument.Sections.Count; i++) {
	d = ActiveDocument.Sections.Item(i);
	if (d.Type == bqDashboard) {
		try {
			if (d.ControlExists("lblDisclaimer")) {
				d.Shapes.RemoveShape("lblDisclaimer");
			}
		//	lTop = d.Shapes["lblReportHeader"].Placement.YOffset + d.Shapes["lblReportHeader"].Placement.Height - 2;
		//	lLeft = d.Shapes["lblReportHeader"].Placement.XOffset + (d.Shapes["lblReportHeader"].Placement.Width / 2) - 165;
		//	s = d.Shapes.CreateShape(bqTextLabel);
		//	s.Name = "lblDisclaimer";
		//	s.Alignment = bqAlignCenter;
		//	with (s.Font) {
		//		Color = 16711680;
		//		Name = "Arial";
		//		Size = 9;
		//		Style = bqFontStyleBold;
		//		Effect = bqFontEffectNone;
		//	}
		//	with (s.Placement) {
		//		XOffset = lLeft;
		//		YOffset = lTop;
		//		Height = 15;
		//		Width = 330;
		//	}
		//	s.Text = "DATA PRESENTED HERE IS FOR TEST PURPOSES ONLY";
		}
		catch (e) {
			Console.Writeln("error placing disclaimer");
		}
	}
	if (d.Type == bqReport) {
		try {
			var oShapes = d.PageHeader.Shapes;
			for (var j = 1; j <= oShapes.Count; j++) {
				if (oShapes.Item(j).Text == "DATA PRESENTED HERE IS FOR TEST PURPOSES ONLY") {
					oShapes.Item(j).Text = "";
				}
			}
			//ActiveDocument.Sections["Expenditure Summary - Org by Fiscal Year"].PageHeader.Shapes["TextLabel2"].Text
		}
		catch (e) { }
	}
}


ActiveDocument.HideTools = function (arr) {
	for (var i = 1; i <= ActiveDocument.Sections.Count; i++) {
		if (ActiveDocument.Sections.Item(i).Type == bqDashboard) {
			for (var j = 0; j < arr.length; j++) {
				try {
					ActiveDocument.Sections.Item(i).Shapes["pic" + arr[j]].Visible = false;
					ActiveDocument.Sections.Item(i).Shapes["pic" + arr[j] + "D"].Visible = false;
				}
				catch (e) {
					//	not found
				}
			}
		}
	}
}
