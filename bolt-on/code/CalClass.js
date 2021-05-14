
Object.extend(Application, {
	Calendar: Class.create({
		initialize:  function(oParent) {
			_logCal = new Application.logger();
			_logCal.Level(_logCal.LEVEL.DEBUG);
			_logCal.Verbose(false);
			
			if (oParent == undefined) {
				//	this is probably an error
				_logCal.error("The parent object of this calendar was not identified.");
				this.Valid = false;
				return;
			}
			else {
				if (oParent.Type == bqDashboard) {
					//	if there is already a calendar object on this dashboard, end
					var sTemp = "";
					try {
						if (oParent.cal.Type == "Calendar" && oParent.cal.Valid) {
							_logCal.error("This dashboard already has a calendar object.");
							this.Valid = false;
							return;
							//break;
						}
					}
					catch (e) {
						//	the object doesn't have a member named cal of type "Calendar"
						//	we're good to go
					}
					_parent = oParent;
				}
				else {
					_logCal.error("The parent object of this calendar is not a dashboard.");
					this.Valid = false;
					return;
				}
			}
			
			_Day = new Array();		//	A 42-member array of DayRect objects
			_DOW = new Array();		//	A 7-member array of DOWRect objects
			_Width = 280;
			_Height = 280;
			_Top = 100;
			_Left = 100;
			_Margin = 10;
			_dayWidth = 30;			//	width of DayRect and DOWRect objects
			_dayHeight = 30;		//	height of DayRect, DOWRect, and Title objects
			_colorToday = parseInt("ffff88", 16);
			_colorSelected = parseInt("88ff88", 16);
			_colorDefault = parseInt("ccffcc", 16);
			_colorBackground = parseInt("ffffff", 16);
			_Visible = false;
			_Format = "mm/dd/yyyy";
			if (typeof sFormat == "string" && sFormat != "") _Format = sFormat;
			_Font = new Application.Font();
			_Font.Wrapper = this;
			//_Font.Size(9);
			_Date = new Date((new Date()).getFullYear(), (new Date()).getMonth(), (new Date()).getDate());
			_Target = new Object();
			
			_Container = _parent.Shapes.CreateShape(bqRectangle);
			_Container.Wrapper = this;
			_Container.Visible = false;
			_Container.Name = this.Name + "_container";
			_Container.SelectDay = function (o) {
				o.Fill.Color = parseInt("88ff88", 16);
				this.SelectedDate = new Date(cal_drpYear.Item(cal_drpYear.SelectedIndex), cal_drpMonth.SelectedIndex - 1, o.Text);
			};
			_Container.Fill.Color = _colorBackground;
			
			_Border = new Application.CalBorder();
			
			_Title = new Application.CalTitle(oParent);
			_Title.Wrapper = this;
			_Title.Year(_Date.getFullYear());
			_Title.Month(_Date.getMonth());
			
			strTempScript = "";
			for (var i = 0; i < 42; i++) {
				_Day.push(new CalDayRect());
				_Day[i].Object = _parent.Shapes.CreateShape(bqTextLabel);
				_Day[i].Object.Visible = false;
				_Day[i].Object.Name = this.Name + "_day_" + i.toString();
				_Day[i].Object.Fill.Color = _colorDefault;
				strTempScript = this.Name + "_obj.SelectedDate(new Date(" + this.Name + "_YearList.GetValue()[0], " + this.Name + "_MonthList.SelectedIndex - 1, this.Text));\r\n";
				strTempScript += this.Name + "_obj.Target().Text = " + this.Name + "_obj.SelectedDate().format(" + this.Name + "_obj.OutputFormat());\r\n";
				strTempScript += this.Name + "_obj.Hide();";
				_Day[i].Object.EventScripts["OnClick"].Script = strTempScript;
				_Day[i].Object.VerticalAlignment = bqAlignMiddle;
				_Day[i].Object.Alignment = bqAlignCenter;
				_Day[i].Object.Wrapper = this;
			}
			
			for (var i = 0; i < 7; i++) {
				_DOW.push(new CalDOWRect());
				_DOW[i].Object = _parent.Shapes.CreateShape(bqTextLabel);
				_DOW[i].Object.Visible = false;
				_DOW[i].Object.Name = this.Name + "_dow_" + i.toString();
				_DOW[i].Object.Text = ActiveDocument.WeekdayAbbrev[i];
				_DOW[i].Object.VerticalAlignment = bqAlignMiddle;
				_DOW[i].Object.Alignment = bqAlignCenter;
				_DOW[i].Object.Wrapper = this;
			}
			
			_parent.cal_obj = this;
			_dayWidth = Math.floor((_Width - (2 * _Margin)) / 7);
			_dayHeight = Math.floor((_Height - (2 * _Margin)) / 8);
			for (var i = 0; i < 6; i++) {
				for (var j = 0; j < 7; j++) {
					_Day[(i * 7) + j].Left(_Left + _Margin + (j * _dayWidth));
					_Day[(i * 7) + j].Top(_Top + _Margin + ((i + 2) * _dayHeight));
					_Day[(i * 7) + j].Width(_dayWidth);
					_Day[(i * 7) + j].Height(_dayHeight);
					if ((new Date(_Title.YearList.GetValue()[0], _Title.MonthList.GetValue()[0], _Day[(i * 7) + j].Object.Text)).toString() == _Date.toString()) {
						_Day[(i * 7) + j].Object.Fill.Color = _colorSelected;
					}
					else if ((new Date(_Title.YearList.GetValue()[0], _Title.MonthList.GetValue()[0], _Day[(i * 7) + j].Object.Text)).toString() == (new Date((new Date()).getFullYear(), (new Date).getMonth(), (new Date).getDate())).toString()) {
						_Day[(i * 7) + j].Object.Fill.Color = _colorToday;
					}
					else {
						_Day[(i * 7) + j].Object.Fill.Color = _colorDefault;
					}
				}
			}
			for (var i = 0; i < 7; i++) {
				_DOW[i].Left(_Left + _Margin + (i * _dayWidth));
				_DOW[i].Top(_Top + _Margin + _dayHeight);
				_DOW[i].Width(_dayWidth);
				_DOW[i].Height(_dayHeight);
			}
			_Width = (7 * _dayWidth) + (2 * _Margin);
			_Height = (8 * _dayHeight) + (2 * _Margin);
			_Container.Placement.XOffset = _Left;
			_Container.Placement.YOffset = _Top;
			_Container.Placement.Width = _Width;
			_Container.Placement.Height = _Height;
			_Container.Fill.Color = _colorBackground;
			_Title.Left(_Left + _Margin);
			_Title.Top(_Top + _Margin);
			_Title.Width(7 * _dayWidth);
			_Title.Height(_dayHeight);
			_Title.MonthList.OnSelection();
			
			this.Valid = true;
			
		},
		
		Type:  "Calendar",
		Name:  "cal",
		
		setupDays: function (intYear, intMonth) {
			var firstday = new Date(intYear, intMonth, 1);
			var DayStart = firstday.getDay();
			
			_logCal.trace(intYear);
			_logCal.trace(intMonth);
			
			//*********************************************
			
			for (i = 0; i < 42; i++) {
				var RectDay = new Date(intYear, intMonth, i - DayStart);
				
				_Day[i].Object.Text = i - DayStart;
				
				if ((new Date(_Title.YearList.GetValue()[0], _Title.MonthList.SelectedIndex - 1, _Day[i].Object.Text)).toString() == _Date.toString()) {
					_Day[i].Object.Fill.Color = _colorSelected;
				}
				else if ((new Date(_Title.YearList.GetValue()[0], _Title.MonthList.SelectedIndex - 1, _Day[i].Object.Text)).toString() == (new Date((new Date()).getFullYear(), (new Date).getMonth(), (new Date).getDate())).toString()) {
					_Day[i].Object.Fill.Color = _colorToday;
				}
				else {
					_Day[i].Object.Fill.Color = _colorDefault;
				}
				
				if (i - DayStart < 0 || RectDay.getMonth() != intMonth || !_Visible) {
					_Day[i].Object.Visible = false;
				}
				else {
					_Day[i].Object.Visible = true;
				}
			}
		},
		
		Parent:  function() {
			return _parent;
		},
		
		Day: function () {
			return _Day;
		},
		
		DOW: function () {
			return _DOW;
		},
		
		Title: function () {
			return _Title;
		},
		
		Border: function () {
			return _Border;
		},
		
		OutputFormat: function (strFormat) {
			if (typeof strFormat == "string" && strFormat != "") {
				_Format = strFormat;
			}
			return _Format;
		},
		
		Target: function (objTarget) {
			if (objTarget == undefined) {
			}
			else if (typeof objTarget == "object") {
				if ((new Array(bqTextLabel, bqTextBox)).contains(objTarget.Type)) {
					_Target = objTarget;
				}
			}
			else if ((new Array(bqTextLabel, bqTextBox)).contains(objTarget.Type)) {
				_Target = objTarget;
			}
			else {
			}
			return _Target;
		},
		
		ColorToday: function (intColor) {
			if (typeof intColor == "number") {
				if (intColor >= 0 && parseInt(intColor, 16) <= parseInt("ffffff", 16)) {
					_colorToday = intColor;
				}
			}
			else if (typeof intColor == "string") {
				if (parseInt(intColor, 16) >= 0 && parseInt(intColor, 16) <= parseInt("ffffff", 16)) {
					_colorToday = parseInt(intColor, 16);
				}
			}
			for (var i = 0; i < 42; i++) {
				if ((new Date(_Title.YearList.GetValue()[0], _Title.MonthList.SelectedIndex - 1, _Day[i].Object.Text)).toString() == _Date.toString()) {
					_Day[i].Object.Fill.Color = _colorSelected;
				}
				else if ((new Date(_Title.YearList.GetValue()[0], _Title.MonthList.SelectedIndex - 1, _Day[i].Object.Text)).toString() == (new Date((new Date()).getFullYear(), (new Date).getMonth(), (new Date).getDate())).toString()) {
					_Day[i].Object.Fill.Color = _colorToday;
				}
				else {
					_Day[i].Object.Fill.Color = _colorDefault;
				}
			}
			return _colorToday;
		},
		
		ColorSelected: function (intColor) {
			if (typeof intColor == "number") {
				if (intColor >= 0 && parseInt(intColor, 16) <= parseInt("ffffff", 16)) {
					_colorSelected = intColor;
				}
			}
			else if (typeof intColor == "string") {
				if (parseInt(intColor, 16) >= 0 && parseInt(intColor, 16) <= parseInt("ffffff", 16)) {
					_colorSelected = parseInt(intColor, 16);
				}
			}
			for (var i = 0; i < 42; i++) {
				if ((new Date(_Title.YearList.GetValue()[0], _Title.MonthList.SelectedIndex - 1, _Day[i].Object.Text)).toString() == _Date.toString()) {
					_Day[i].Object.Fill.Color = _colorSelected;
				}
				else if ((new Date(_Title.YearList.GetValue()[0], _Title.MonthList.SelectedIndex - 1, _Day[i].Object.Text)).toString() == (new Date((new Date()).getFullYear(), (new Date).getMonth(), (new Date).getDate())).toString()) {
					_Day[i].Object.Fill.Color = _colorToday;
				}
				else {
					_Day[i].Object.Fill.Color = _colorDefault;
				}
			}
			return _colorSelected;
		},
		
		ColorDefault: function (intColor) {
			if (typeof intColor == "number") {
				if (intColor >= 0 && parseInt(intColor, 16) <= parseInt("ffffff", 16)) {
					_colorDefault = intColor;
				}
			}
			else if (typeof intColor == "string") {
				if (parseInt(intColor, 16) >= 0 && parseInt(intColor, 16) <= parseInt("ffffff", 16)) {
					_colorDefault = parseInt(intColor, 16);
				}
			}
			
			for (var i = 0; i < 42; i++) {
				if ((new Date(_Title.YearList.GetValue()[0], _Title.MonthList.SelectedIndex - 1, _Day[i].Object.Text)).toString() == _Date.toString()) {
					_Day[i].Object.Fill.Color = _colorSelected;
				}
				else if ((new Date(_Title.MonthList.GetValue()[0] + " " + _Day[i].Object.Text + ", " + _Title.YearList.GetValue()[0])).toString() == (new Date((new Date()).getFullYear(), (new Date).getMonth(), (new Date).getDate())).toString()) {
					_Day[i].Object.Fill.Color = _colorToday;
				}
				else {
					_Day[i].Object.Fill.Color = _colorDefault;
				}
			}
			return _colorDefault;
		},
		
		ColorBackground: function (intColor) {
			if (typeof intColor == "number") {
				if (intColor >= 0 && parseInt(intColor, 16) <= parseInt("ffffff", 16)) {
					_colorBackground = intColor;
				}
			}
			else if (typeof intColor == "string") {
				if (parseInt(intColor, 16) >= 0 && parseInt(intColor, 16) <= parseInt("ffffff", 16)) {
					_colorBackground = parseInt(intColor, 16);
				}
			}
			_Container.Fill.Color = _colorBackground;
			return _colorBackground;
		},
		
		SelectedDate: function (dteDate) {
			if (Date.prototype.isPrototypeOf(dteDate)) {
				_Date = dteDate;
				_Title.Month(_Date.getMonth());
				_Title.Year(_Date.getFullYear());
				_Title.MonthList.OnSelection();
			}
			return _Date;
		},
		
		Show: function () {
			for (var i = 1; i < _parent.Shapes.Count; i++) {
				if (_parent.Shapes.Item(i).Name.substr(0, 4) == "cal_") {
					_Visible = true;
					_parent.Shapes.Item(i).Visible = true;
				}
			}
			this.setupDays(_Title.Year(), _Title.Month());
		},
		
		Hide: function () {
			for (var i = 1; i < _parent.Shapes.Count; i++) {
				if (_parent.Shapes.Item(i).Name.substr(0, 4) == "cal_") {
					_Visible = false;
					_parent.Shapes.Item(i).Visible = false;
				}
			}
		},
		
		Kill: function () {
			var arrShapes = new Array();
			for (var i = 1; i < _parent.Shapes.Count; i++) {
				if (_parent.Shapes.Item(i).Name.substr(0, 4) == "cal_") {
					arrShapes.push(_parent.Shapes.Item(i).Name);
				}
			}
			for (var i = 0; i < arrShapes.length; i++) {
				_parent.Shapes.RemoveShape(arrShapes[i]);
			}
			_parent.cal_obj = null;
			this.Valid = false;
		},
		
		Margin: function (intMargin) {
			if (typeof intMargin == "number") {
				_Margin = intMargin;
				intTemp = _Width - (2 * _Margin);
				_dayWidth = Math.floor(intTemp / 7);
				intTemp = _Height - (2 * _Margin);
				_dayHeight = Math.floor(intTemp / 8);
				for (var i = 0; i < 6; i++) {
					for (var j = 0; j < 7; j++) {
						_Day[(i * 7) + j].Width(_dayWidth);
						_Day[(i * 7) + j].Left(_Left + _Margin + (j * _dayWidth));
						_Day[(i * 7) + j].Height(_dayHeight);
						_Day[(i * 7) + j].Top(_Top + _Margin + ((i + 2) * _dayHeight));
					}
				}
				for (var i = 0; i < 7; i++) {
					_DOW[i].Width(_dayWidth);
					_DOW[i].Left(_Left + _Margin + (i * _dayWidth));
					_DOW[i].Height(_dayHeight);
					_DOW[i].Top(_Top + _Margin + _dayHeight);
				}
				_Width = (7 * _dayWidth) + (2 * _Margin);
				_Height = (8 * _dayHeight) + (2 * _Margin);
				_Container.Placement.Width = _Width;
				_Container.Placement.Height = _Height;
				_Title.Width(7 * _dayWidth);
				_Title.Height(_dayHeight);
				_Title.Left(_Left + _Margin);
				_Title.Top(_Top + _Margin);
			}
			return _Margin;
		},
		
		Width: function (intWidth) {
			if (typeof intWidth == "number") {
				intTemp = intWidth - (2 * _Margin);
				_dayWidth = Math.floor(intTemp / 7);
				for (var i = 0; i < 6; i++) {
					for (var j = 0; j < 7; j++) {
						_Day[(i * 7) + j].Width(_dayWidth);
						_Day[(i * 7) + j].Left(_Left + _Margin + (j * _dayWidth));
					}
				}
				for (var i = 0; i < 7; i++) {
					_DOW[i].Width(_dayWidth);
					_DOW[i].Left(_Left + _Margin + (i * _dayWidth));
				}
				_Title.Width(7 * _dayWidth);
				_Width = (7 * _dayWidth) + (2 * _Margin);
				_Container.Placement.Width = _Width;
			}
			return _Width;
		},
		
		Height: function (intHeight) {
			if (typeof intHeight == "number") {
				intTemp = intHeight - (2 * _Margin);
				_dayHeight = Math.floor(intTemp / 8);
				for (var i = 0; i < 6; i++) {
					for (var j = 0; j < 7; j++) {
						_Day[(i * 7) + j].Height(_dayHeight);
						_Day[(i * 7) + j].Top(_Top + _Margin + ((i + 2) * _dayHeight));
					}
				}
				for (var i = 0; i < 7; i++) {
					_DOW[i].Height(_dayHeight);
					_DOW[i].Top(_Top + _Margin + _dayHeight);
				}
				_Title.Height(_dayHeight);
				_Height = (8 * _dayHeight) + (2 * _Margin);
				_Container.Placement.Height = _Height;
			}
			return _Height;
		},
		
		Top: function (intTop) {
			if (typeof intTop == "number") {
				_Top = intTop;
				for (var i = 0; i < 6; i++) {
					for (var j = 0; j < 7; j++) {
						_Day[(i * 7) + j].Top(_Top + _Margin + ((i + 2) * _dayHeight));
					}
				}
				for (var i = 0; i < 7; i++) {
					_DOW[i].Top(_Top + _Margin + _dayHeight);
				}
				_Title.Top(_Top + _Margin);
				_Container.Placement.YOffset = _Top;
			}
			return _Top;
		},
		
		Left: function (intLeft) {
			if (typeof intLeft == "number") {
				_Left = intLeft;
				for (var i = 0; i < 6; i++) {
					for (var j = 0; j < 7; j++) {
						_Day[(i * 7) + j].Left(_Left + _Margin + (j * _dayWidth));
					}
				}
				for (var i = 0; i < 7; i++) {
					_DOW[i].Left(_Left + _Margin + (i * _dayWidth));
				}
				_Title.Left(_Left + _Margin);
				_Container.Placement.XOffset = _Left;
			}
			return _Left;
		},
		
		Font: function (oFont) {
			if (typeof oFont == "object") {
				_Font = oFont;
				
				for (var i = 0; i < 42; i++) {
					_Day[i].Object.Font.Name = _Font.Name();
					_Day[i].Object.Font.Style = _Font.Style();
					_Day[i].Object.Font.Effect = _Font.Effect();
					_Day[i].Object.Font.Size = _Font.Size();
					_Day[i].Object.Font.Color = _Font.Color();
				}
				
				for (var i = 0; i < 7; i++) {
					_DOW[i].Object.Font.Name = _Font.Name();
					_DOW[i].Object.Font.Style = _Font.Style();
					_DOW[i].Object.Font.Effect = _Font.Effect();
					_DOW[i].Object.Font.Size = _Font.Size();
					_DOW[i].Object.Font.Color = _Font.Color();
				}
				
				_Title.MonthLabel.Font.Name = _Font.Name();
				_Title.MonthLabel.Font.Style = _Font.Style();
				_Title.MonthLabel.Font.Effect = _Font.Effect();
				_Title.MonthLabel.Font.Size = _Font.Size();
				_Title.MonthLabel.Font.Color = _Font.Color();
				
				_Title.YearLabel.Font.Name = _Font.Name();
				_Title.YearLabel.Font.Style = _Font.Style();
				_Title.YearLabel.Font.Effect = _Font.Effect();
				_Title.YearLabel.Font.Size = _Font.Size();
				_Title.YearLabel.Font.Color = _Font.Color();
				
				_Title.MonthList.Font.Name = _Font.Name();
				_Title.MonthList.Font.Style = _Font.Style();
				_Title.MonthList.Font.Effect = _Font.Effect();
				_Title.MonthList.Font.Size = _Font.Size();
				
				_Title.YearList.Font.Name = _Font.Name();
				_Title.YearList.Font.Style = _Font.Style();
				_Title.YearList.Font.Effect = _Font.Effect();
				_Title.YearList.Font.Size = _Font.Size();
				DoEvents()
			}
			return _Font;
		},
		
		toString: function () {
			return "[object Calendar]";
		}
	})
});



Application.CalTitle = function (oDashboard) {
	//Console.Writeln("Entering CalTitle");
	//_Font.Wrapper = this;
	var _Month = 0;
	var _Year = (new Date()).getFullYear();
	var _Left = 0;
	var _Top = 0;
	var _Width = 0;
	var _Height = 0;
	var arrYears = new Array();
	
	this.MonthList = oDashboard.Shapes.CreateShape(bqDropDown);
	this.MonthList.Visible = false;
	this.MonthList.Name = "cal_MonthList";
	this.MonthList.EventScripts["OnSelection"].Script = "cal_obj.setupDays(cal_YearList.Item(cal_YearList.SelectedIndex), cal_MonthList.SelectedIndex - 1);";
	this.MonthList.Wrapper = this;
	this.YearList = oDashboard.Shapes.CreateShape(bqDropDown);
	this.YearList.Visible = false;
	this.YearList.Name = "cal_YearList";
	this.YearList.EventScripts["OnSelection"].Script = "cal_obj.setupDays(cal_YearList.Item(cal_YearList.SelectedIndex), cal_MonthList.SelectedIndex - 1);";
	this.YearList.Wrapper = this;
	this.MonthLabel = oDashboard.Shapes.CreateShape(bqTextLabel);
	this.MonthLabel.Visible = false;
	this.MonthLabel.VerticalAlignment = bqAlignTop;
	this.MonthLabel.Name = "cal_title_month";
	this.MonthLabel.Text = "Month:";
	this.YearLabel = oDashboard.Shapes.CreateShape(bqTextLabel);
	this.YearLabel.Visible = false;
	this.YearLabel.VerticalAlignment = bqAlignTop;
	this.YearLabel.Name = "cal_title_year";
	this.YearLabel.Text = "Year:";
	
    //this.Font = function (vFont) {
	//	return _Font;
	//};
	
	this.Month = function (intText) {
		if (typeof intText == "number") {
			if (intText > -1 && intText < 12) {
				_Month = intText;
				this.MonthList.Select(intText + 1);
				//this.MonthList.OnSelection();
			}
		}
		return _Month;
	};
	
	this.Year = function (intYear) {
		if (typeof intText == "Number") {
			_Year = intYear;
			this.YearList.SetValue(new Array(_Year));
			this.YearList.OnSelection();
		}
		return _Year;
	};
	
	this.MonthName = function () {
		return ActiveDocument.MonthName[_Month];
	};
	
	this.Top = function (intTop) {
		if (typeof intTop == "number") {
			_Top = intTop;
			this.YearLabel.Placement.YOffset = _Top;
			this.YearList.Placement.YOffset = _Top;
			this.MonthLabel.Placement.YOffset = _Top;
			this.MonthList.Placement.YOffset = _Top;
		}
		return _Top;
	};
	
	this.Left = function (intLeft) {
		if (typeof intLeft == "number") {
			_Left = intLeft;
			this.YearLabel.Placement.XOffset = _Left;
			this.YearList.Placement.XOffset = _Left + (_Width / 6);
			this.MonthLabel.Placement.XOffset = _Left + (_Width / 2);
			this.MonthList.Placement.XOffset = _Left + (_Width * 2 / 3);
		}
		return _Left;
	};
	
	this.Height = function (intHeight) {
		if (typeof intHeight == "number") {
			_Height = intHeight;
			this.YearLabel.Placement.Height = _Height;
			this.YearList.Placement.Height = _Height;
			this.MonthLabel.Placement.Height = _Height;
			this.MonthList.Placement.Height = _Height;
		}
		return _Height;
	};
	
	this.Width = function (intWidth) {
		if (typeof intWidth == "number") {
			_Width = intWidth;
			this.YearLabel.Placement.Width = _Width / 6;
			this.YearList.Placement.Width = _Width / 3;
			this.MonthLabel.Placement.Width = _Width / 6;
			this.MonthList.Placement.Width = _Width / 3;
			this.YearLabel.Placement.XOffset = _Left;
			this.YearList.Placement.XOffset = _Left + (_Width / 6);
			this.MonthLabel.Placement.XOffset = _Left + (_Width / 2);
			this.MonthList.Placement.XOffset = _Left + (_Width * 2 / 3);
		}
		return _Width;
	};
	
	for (var i = 1995; i < 2039; i++) {
		arrYears.push(i.toString());
	}
	
	ActiveDocument.loadObjectMethods();
	
	this.YearList.SetList(arrYears);
	this.MonthList.SetList(ActiveDocument.MonthName);
	this.YearList.SetValue(new Array((new Date()).getFullYear().toString()));
	this.MonthList.SetValue(new Array(ActiveDocument.MonthName[(new Date()).getMonth()]));
}



Application.CalBorder = function () {
	//Console.Writeln("Entering CalBorder");
    //DashStyle
	var _dashStyleDash = 3;
	var _dashStyleDot = 2;
	var _dashStyleDotDash = 4;
	var _dashStyleDotDotDash = 5;
	var _dashStyleSolid = 1;
	
	//var _Style = 0;		//	0 = flat, 1 = raised, 2 = sunken
	var _Size = 1;	//	flat style only, use 0 for no border
	var _Color = parseInt("000000", 16);
	var _DashStyle = _dashStyleSolid;
	
// TextLabel.Line.Color
// TextLabel.Line.DashStyle
// TextLabel.Line.Width
	
	//this.Styles = function () {
	//	this.Flat = function () {
	//		return 0;
	//	}
	//	this.Raised = function () {
	//		return 1;
	//	}
	//	this.Sunken = function () {
	//		return 2;
	//	}
	//};
	
	this.DashStyles = function () {
		this.Dash = function () {
			return _dashStyleDash;
		};
		this.Dot = function () {
			return _dashStyleDot;
		};
		this.DotDash = function () {
			return _dashStyleDotDash;
		};
		this.DotDotDash = function () {
			return _dashStyleDotDotDash;
		};
		this.Solid = function () {
			return _dashStyleSolid;
		};
	};
	
	//this.Style = function (intStyle) {
	//	if (typeof intStyle == "number") {
	//		_Style = intStyle;
	//	}
	//	return _Style;
	//};
	
	this.Size = function (intSize) {
		if (typeof intSize == "number") {
			_Size = intSize;
			this.Wrapper.Object.Line.Width = _Size;
		}
		return _Size;
	};
	
	this.Color = function (intColor) {
		if (typeof intColor == "number") {
			if (intColor >= 0 && intColor <= parseInt("ffffff", 16)) {
				_Color = intColor;
				this.Wrapper.Object.Line.Color = _Color;
			}
		}
		else if (typeof intColor == "string") {
			if (parseInt(intColor, 16) >= 0 && parseInt(intColor, 16) <= parseInt("ffffff", 16)) {
				_Color = parseInt(intColor, 16);
				this.Wrapper.Object.Line.Color = _Color;
			}
		}
		return _Color;
	};
	
	this.DashStyle = function (intDashStyle) {
		if (typeof intDashStyle == "number") {
			_DashStyle = intDashStyle;
			this.Wrapper.Object.Line.DashStyle = _DashStyle;
		}
		return _DashStyle;
	};
}


Application.CalDayRect = function () {
	//Console.Writeln("Entering CalDayRect");
	var _Top = 0;
	var _Left = 0;
	var _Height = 30;
	var _Width = 30;
	var _BackColor = parseInt("ffffff", 16);
	var _Value = 0;
	//_Font.Wrapper = this;
	var _Border = new CalBorder();
	_Border.Wrapper = this;
	
	this.Object = new Object();
	
// TextLabel.Line.Color
// TextLabel.Line.DashStyle
// TextLabel.Line.Width
// TextLabel.Fill.Color
// TextLabel.Fill.Pattern
// TextLabel.EventScripts["OnClick"].Script
// TextLabel.VerticalAlignment
// TextLabel.Alignment
// TextLabel.Font.Color
// TextLabel.Font.Effect
// TextLabel.Font.Name
// TextLabel.Font.Size
// TextLabel.Font.Style
// TextLabel.Text
// TextLabel.TextWrap
// TextLabel.Visible
// TextLabel.Comments
// TextLabel.Placement.YOffset
// TextLabel.Placement.XOffset
// TextLabel.Placement.Height
// TextLabel.Placement.Width
	
	this.Top = function (intTop) {
		if (typeof intTop == "number") {
			_Top = intTop;
			this.Object.Placement.YOffset = _Top;
		}
		return _Top;
	};
	
	this.Left = function (intLeft) {
		if (typeof intLeft == "number") {
			_Left = intLeft;
			this.Object.Placement.XOffset = _Left;
		}
		return _Left;
	};
	
	this.Height = function (intHeight) {
		if (typeof intHeight == "number") {
			_Height = intHeight;
			this.Object.Placement.Height = _Height;
		}
		return _Height;
	};
	
	this.Width = function (intWidth) {
		if (typeof intWidth == "number") {
			_Width = intWidth;
			this.Object.Placement.Width = _Width;
		}
		return _Width;
	};
	
	this.BackColor = function (intBackColor) {
		if (typeof intBackColor == "number") {
			_BackColor = intBackColor;
			this.Object.Fill.Color = _BackColor;
		}
		else if (typeof intBackColor == "string") {
			if (parseInt(intBackColor, 16) >= 0 && parseInt(intBackColor, 16) <= parseInt("ffffff", 16)) {
				_BackColor = parseInt(intBackColor, 16);
				this.Object.Fill.Color = _BackColor;
			}
		}
		return _BackColor;
	};
	
	this.Value = function (intValue) {
		if (typeof intValue == "number") {
			_Value = intValue;
			this.Object.Text = _Value.toString();
		}
		return _Value;
	};
	
	//this.Font = function () {
	//	return _Font;
	//};
	
	this.Border = function () {
		return _Border;
	};
}


Application.CalDOWRect = function () {
	//Console.Writeln("Entering CalDOWRect");
	var _Top = 0;
	var _Left = 0;
	var _Height = 30;
	var _Width = 30;
	var _BackColor = parseInt("ffffff", 16);
	var _Value = 0;
	var _Label = "Sun";
	var _Border = new CalBorder();
	
	this.Object = new Object();
	
	this.Top = function (intTop) {
		if (typeof intTop == "number") {
			_Top = intTop;
			this.Object.Placement.YOffset = _Top;
		}
		return _Top;
	};
	
	this.Left = function (intLeft) {
		if (typeof intLeft == "number") {
			_Left = intLeft;
			this.Object.Placement.XOffset = _Left;
		}
		return _Left;
	};
	
	this.Height = function (intHeight) {
		if (typeof intHeight == "number") {
			_Height = intHeight;
			this.Object.Placement.Height = _Height;
		}
		return _Height;
	};
	
	this.Width = function (intWidth) {
		if (typeof intWidth == "number") {
			_Width = intWidth;
			this.Object.Placement.Width = _Width;
		}
		return _Width;
	};
	
	this.BackColor = function (intBackColor) {
		if (typeof intBackColor == "number") {
			_BackColor = intBackColor;
			this.Object.Fill.Color = _BackColor;
		}
		else if (typeof intBackColor == "string") {
			if (parseInt(intBackColor, 16) >= 0 && parseInt(intBackColor, 16) <= parseInt("ffffff", 16)) {
				_BackColor = parseInt(intBackColor, 16);
			}
		}
		return _BackColor;
	};
	
	this.Value = function (intValue) {
		if (typeof intValue == "number") {
			if (intValue > -1 && intValue < 7) {
				_Value = intValue;
				_Label = ActiveDocument.WeekdayAbbrev[intValue];
				this.Object.Text = _Label;
			}
		}
		return _Value;
	};
	
	//this.Font = function () {
	//	return _Font;
	//};
	
	this.Border = function () {
		return _Border;
	};
}
