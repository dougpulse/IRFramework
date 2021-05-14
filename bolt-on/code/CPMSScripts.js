//	File:		FinancialScripts.js
//	Language:	JavaScript
//	Purpose:	to provide common scripts that apply to the user interfaces that have been 
//				developed for the Financial Data Mart but are not common to all files using scripting
//	Author:		doug pulse
//	Revision History:
//	doug pulse		5/20/2010	inaugural
//

ActiveDocument.RunQueries = function (arrQuery, arrListName, arrLimitName, dash, o, pb) {
	//	Process:	RunQuery()
	//	Language:	JavaScript
	//	Purpose:	Set limits and run each query
	//	Inputs:		arrQuery		Array containing the Query sections to process
	//				arrListName		Array containing the names of the objects containing the limit criteria
	//				arrLimitName	Array containing the names of the limits
	//				o				the object calling this function
	//				dash			parent dashboard of the object calling this function
	//				pb				(optional) progressbar to use
	//	

	//ProgressBar = function (dashboard, left, top, width, height, value, maxvalue)
	//dash.GoProgress = new wsdot_ProgressBar(dash, dash.lblCriteria.Placement.XOffset, dash.lblCriteria.Placement.YOffset + o.Placement.Height, dash.lblCriteria.Placement.Width, o.Placement.Height, 0, 100);
	//dash.GoProgress = new wsdot_ProgressBar(dash, dash.lblCriteria.Placement.XOffset, o.Placement.YOffset + o.Placement.Height + 8, dash.lblCriteria.Placement.Width, o.Placement.Height, 0, 100);
	if (typeof pb != "undefined") {
		pb.Show();
	}
	
	
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
							arr = ActiveDocument.Sections["Selection Criteria"].Shapes["txt" + arrListName[i]].Text.split(",");
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
		
		arrQuery[k].Process();
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
	
	var t = ActiveDocument.Sections["Selection Criteria"].Shapes["txtVol"];			//	textbox holding filter values
	var o = ActiveDocument.Sections["Selection Criteria"].Shapes["drpOperVol"];		//	operator dropdown
	
	var q = new Array();		//	limits to use to take advantage of partitioning
	for (var i = 0; i < ActiveDocument.arrQuery.length; i++) {
		//Console.Writeln(ActiveDocument.arrQuery[i].Name);
		if (ActiveDocument.arrQueryPart[i]) q.push(ActiveDocument.arrQuery[i].Limits["Capitalprogramvolumeid"]);
	}
	
	var d = ActiveDocument.Sections["r-Capital Program Volume"];
	var l = d.Limits["Capital Program Volume"];		//	limit objects to use in [r-Capital Program Volume]
	
	//	adjust filters to take advantage of partitioning
	//	don't ignore, and remove all filter values from the partition filter
	for (var i = 0; i < q.length; i++) {
		q[i].Ignore = false;
		q[i].CustomValues.RemoveAll();
		q[i].SelectedValues.RemoveAll();
	}
	
	if (t.Text.length != 0) {
		s = t.Text.split(",");
		l.Ignore = false;
		//Console.Writeln(l.Operator = o.SelectedIndex);
		l.CustomValues.RemoveAll();
		l.SelectedValues.RemoveAll();
		for (var j = 0; j < s.length; j++) {
			l.CustomValues.Add(s[j]);
			l.SelectedValues.Add(s[j]);
		}
		n++;
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
					q[j].CustomValues.Add(d.Columns["Capitalprogramvolumeid"].GetCell(k));
					q[j].SelectedValues.Add(d.Columns["Capitalprogramvolumeid"].GetCell(k));
				}
			}
		}
	}
	
	l.Ignore = true;
}



ActiveDocument.ExportData = function (oDash, sResultPrefix) {
	//	Purpose:  export the data from the report to pdf or from the results to excel
	var d = new Date();
	var dfile = new String("");
	var iOutFormat, sSection;
	var sBaseName = oDash.Name;
	
	if (oDash.ActiveTab) {
		sBaseName = oDash.ActiveTab;
	}
	
	dfile = g_strUserTempPath + "\\" + sBaseName + "_" + d.format("yyyymmdd");
	
	var iOut = Alert("What do you want to export?", "Work Order", "Printable Document", "Data to Spreadsheet", "Cancel");
	switch (iOut) {
		case 1 :
			dfile += ".pdf";
			iOutFormat = bqExportFormatPDF;
			sSection = sBaseName.replace(/ /gi, "  ");
			ActiveDocument.Sections[sSection].Export(dfile, iOutFormat, true, false);
			Shell(g_strWinPath + "\\explorer.exe", dfile);
			break;
		case 2 :
			dfile += ".mht";
			iOutFormat = bqExportFormatOfficeMHTML;
			sSection = sResultPrefix + "-" + sBaseName;
			ActiveDocument.Sections[sSection].Export(dfile, iOutFormat, true, false);
			var objExcel = new JOOLEObject("Excel.Application");
			objExcel.Visible = true;
			var objWorkbook = objExcel.Workbooks.Open(dfile);
			break;
		case 3 : 
			//	user cancelled
			break;
		default :
			break;
	}
}


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
if (!ActiveDocument.blnCleared) {
	for (var i = 0; i < arrQuery.length; i++) {
		ActiveDocument.arrQuery[i].Limits["Id"].Ignore = false;
		ActiveDocument.arrQuery[i].Process();
		ActiveDocument.arrQuery[i].Limits["Id"].Ignore = true;
	}
}


ActiveDocument.wsdot_GlobalStartup.loadObjectMethods();
//ActiveDocument.Sections["Selection Criteria"].Shapes["btnSetCleanScript"].OnClick();

//	if the file is loaded in the web client, don't prompt to save when the document is closed
if ((Application.Type == bqAppTypePlugInClient) || (Application.Type == bqAppTypeThinClient))
	ActiveDocument.PromptToSave = false;


//	if the file is loaded in the desktop client, make sure all queries are connected to the database
//if (Application.Type == bqAppTypeDesktopClient) {
	var SecCount = ActiveDocument.Sections.Count;
	for (j = 1; j <= SecCount ; j++) {
		if (ActiveDocument.Sections[j].Type == bqQuery) {
			if (ActiveDocument.Sections[j].DataModel.Topics.Count > 0 || ActiveDocument.Sections[j].DataModel.DerivedTables.Count > 0) {
				with (ActiveDocument.Sections[j].DataModel.Connection) {
					Connect();
				}
			}
		}
	}
//}


/*
//	process the queries needed to take advantage of partitioning
ActiveDocument.Sections["q-Partitions"].Process();
ActiveDocument.Sections["q-Fiscal Month"].Process();


//	get partition information
ActiveDocument.dctPartition = new wsdot_Dictionary();
for (var i = 1, s = "", a = new Array(); i <= ActiveDocument.Sections["r-Partitions"].RowCount; i++) {
	s = ActiveDocument.Sections["r-Partitions"].Columns["Databasename"].GetCell(i);
	s += ".";
	s += ActiveDocument.Sections["r-Partitions"].Columns["Facttablename"].GetCell(i);
	a.push(ActiveDocument.Sections["r-Partitions"].Columns["Factcolumnname"].GetCell(i));
	a.push(ActiveDocument.Sections["r-Partitions"].Columns["Dimensiontablename"].GetCell(i));
	a.push(ActiveDocument.Sections["r-Partitions"].Columns["Dimensioncolumnname"].GetCell(i));
	dctPartition.Add(s, a);
}


//	create month arrays to use for monthname dropdown
ActiveDocument.arrCalMonths = ActiveDocument.MonthName.concat("Month99").concat("Month25");
ActiveDocument.arrFY1Months = ActiveDocument.MonthName.concat("Month99").sort(SortFY1MonthNames);
ActiveDocument.arrFY2Months = ActiveDocument.MonthName.concat("Month25").sort(SortFY2MonthNames);


//	set the default value for the biennium criteria [to the current biennium]
var a = new Array();
var blnBien = false;
try { blnBien = (ActiveDocument.Sections["Selection Criteria"].lstBien.Type == bqListBox) }
catch (e) {
	//	lstBien is not a ListBox -- do nothing
}

if (blnBien) {
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
	for (var i = 1; i <= ActiveDocument.Sections["Selection Criteria"].lstYear.SelectedList.Count; i++) {
		a.push(ActiveDocument.Sections["Selection Criteria"].lstYear.SelectedList.ItemIndex(i));
	}

	for (var i = 0; i <= a.length; i++) {
		ActiveDocument.Sections["Selection Criteria"].lstYear.Unselect(a[i]);
	}

	ActiveDocument.Sections["Selection Criteria"].lstYear.Select((new Date()).getFiscalYear() - ActiveDocument.Sections["Selection Criteria"].lstYear.Item(1) + 1);
	try {
		ActiveDocument.Sections["Selection Criteria"].txtYear.Text = (new Date()).getFiscalYear();
	}
	catch (e) {
		//	meaningless error
	}
}
*/
