
 /*  IRScripts
 *   Prototype-style Hyperion IR JavaScript framework, version 0.0.1.0
 *   2010 Doug Pulse
 *
 *--------------------------------------------------------------------------*/
 /*  License
	Source:	http://dev.rubyonrails.org/browser/spinoffs/prototype/trunk/LICENSE?format=raw
	Date:	8/26/2010
	Text:
	Permission is hereby granted, free of charge, to any person obtaining a copy
	of this software and associated documentation files (the "Software"), to deal
	in the Software without restriction, including without limitation the rights
	to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
	copies of the Software, and to permit persons to whom the Software is
	furnished to do so, subject to the following conditions:
	
	THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
	IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
	FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
	AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
	LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
	OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
	SOFTWARE.
 *--------------------------------------------------------------------------*/

Application.Hyperion = {
	Version:  Application.Version,
	
	Type: {
		Insight:   (Application.Type == bqAppTypePlugInClient),
		Explorer:  (Application.Type == bqAppTypeDesktopClient),
		Scheduler: (Application.Type == bqAppTypeScheduler),
		Other:     (Application.Type != bqAppTypePlugInClient && Application.Type != bqAppTypeDesktopClient && Application.Type != bqAppTypeScheduler),
		toString:  function () {
			var sTemp = "";
			switch (true) {
				case this.Insight :
					sTemp = "Insight";
					break;
				case this.Explorer :
					sTemp = "Explorer";
					break;
				case this.Scheduler :
					sTemp = "Scheduler";
					break;
				default :
					sTemp = "Other";
					break;
			}
			return sTemp;
		}
	},
	
	emptyFunction: function() { },
	K: function(x) { return x },
	JSONFilter: /^\/\*-secure-([\s\S]*)\*\/\s*$/
//	ActiveDocument:  Application.ActiveDocument,
//	Documents:  Application.Documents,
//	Toolbars:  Application.Toolbars
};


/* Based on Alex Arnell's inheritance implementation. */
Application.Class = {
	create: function() {
		var parent = null, properties = $A(arguments);
		if (Object.isFunction(properties[0]))
		parent = properties.shift();
		
		function klass() {
			this.initialize.apply(this, arguments);
		}
		
		Object.extend(klass, Class.Methods);
		klass.superclass = parent;
		klass.subclasses = [];
		
		if (parent) {
			var subclass = function() { };
			subclass.prototype = parent.prototype;
			klass.prototype = new subclass;
			parent.subclasses.push(klass);
		}
		
		for (var i = 0; i < properties.length; i++)
			klass.addMethods(properties[i]);
		
		if (!klass.prototype.initialize)
			klass.prototype.initialize = Hyperion.emptyFunction;
		
		klass.prototype.constructor = klass;
		
		return klass;
	}
};

Class.Methods = {
	addMethods: function(source) {
		var ancestor   = this.superclass && this.superclass.prototype;
		var properties = Object.keys(source);
		
		if (!Object.keys({ toString: true }).length)
			properties.push("toString", "valueOf");
		
		for (var i = 0, length = properties.length; i < length; i++) {
			var property = properties[i], value = source[property];
			if (ancestor && Object.isFunction(value) &&
			value.argumentNames().first() == "$super") {
				var method = value, value = Object.extend((function(m) {
					return function() { return ancestor[m].apply(this, arguments) };
				})(property).wrap(method), {
					valueOf:  function() { return method },
					toString: function() { return method.toString() }
				});
			}
			this.prototype[property] = value;
		}
		
		return this;
	}
};

Application.Abstract = { };

Object.extend = function(destination, source) {
	for (var property in source)
		destination[property] = source[property];
	return destination;
};

Object.extend(Object, {
	inspect: function(object) {
		try {
			if (Object.isUndefined(object)) return 'undefined';
			if (object === null) return 'null';
			return object.inspect ? object.inspect() : String(object);
		} catch (e) {
			if (e instanceof RangeError) return '...';
			throw e;
		}
	},
	
	toJSON: function(object) {
		var type = typeof object;
		switch (type) {
			case 'undefined':
			case 'function':
			case 'unknown': return;
			case 'boolean': return object.toString();
		}
		
		if (object === null) return 'null';
		if (object.toJSON) return object.toJSON();
		if (Object.isElement(object)) return;
		
		var results = [];
		for (var property in object) {
			var value = Object.toJSON(object[property]);
			if (!Object.isUndefined(value))
			results.push(property.toJSON() + ': ' + value);
		}
		
		return '{' + results.join(', ') + '}';
	},
	
	toQueryString: function(object) {
		return $H(object).toQueryString();
	},
	
	toHTML: function(object) {
		return object && object.toHTML ? object.toHTML() : String.interpret(object);
	},
	
	keys: function(object) {
		var keys = [];
		for (var property in object)
		keys.push(property);
		return keys;
	},
	
	values: function(object) {
		var values = [];
		for (var property in object)
		values.push(object[property]);
		return values;
	},
	
	clone: function(object) {
		return Object.extend({ }, object);
	},
	
	isElement: function(object) {
		return object && object.nodeType == 1;
	},
	
	isArray: function(object) {
		return object != null && typeof object == "object" && 'splice' in object && 'join' in object;
	},
	
	isHash: function(object) {
		return object instanceof Hash;
	},
	
	isFunction: function(object) {
		return typeof object == "function";
	},
	
	isString: function(object) {
		return typeof object == "string";
	},
	
	isNumber: function(object) {
		return typeof object == "number";
	},
	
	isUndefined: function(object) {
		return typeof object == "undefined";
	},
	
	isDate: function (object) {
		return (new Date(object)).toString() != "Invalid Date";
	},
	
	isInteger: function (object) {
		return (object.toString().search(/^-?[0-9]+$/) == 0);
	},
	
	isUnsignedInteger: function (object) {
		return (object.toString().search(/^[0-9]+$/) == 0);
	}
});

Application.$A = function (iterable) {
	if (!iterable) return [];
	if (iterable.toArray) return iterable.toArray();
	var length = iterable.length || 0, results = new Array(length);
	while (length--) results[length] = iterable[length];
	return results;
}


Object.extend(Function.prototype, {
	argumentNames: function() {
		var names = this.toString().match(/^[\s\(]*function[^(]*\((.*?)\)/)[1].split(",").invoke("strip");
		return names.length == 1 && !names[0] ? [] : names;
	},
	
	bind: function() {
		if (arguments.length < 2 && Object.isUndefined(arguments[0])) return this;
		var __method = this, args = $A(arguments), object = args.shift();
		return function() {
			return __method.apply(object, args.concat($A(arguments)));
		}
	},
	
	//bindAsEventListener: function() {
	//	var __method = this, args = $A(arguments), object = args.shift();
	//	return function(event) {
	//		return __method.apply(object, [event || window.event].concat(args));
	//	}
	//},
	
	curry: function() {
		if (!arguments.length) return this;
		var __method = this, args = $A(arguments);
		return function() {
			return __method.apply(this, args.concat($A(arguments)));
		}
	},
	
	//delay: function() {
	//	var __method = this, args = $A(arguments), timeout = args.shift() * 1000;
	//	return window.setTimeout(function() {
	//		return __method.apply(__method, args);
	//	}, timeout);
	//},
	
	wrap: function(wrapper) {
	var __method = this;
		return function() {
			return wrapper.apply(this, [__method.bind(this)].concat($A(arguments)));
		}
	},
	
	methodize: function() {
		if (this._methodized) return this._methodized;
		var __method = this;
		return this._methodized = function() {
			return __method.apply(null, [this].concat($A(arguments)));
		};
	}
});

//Function.prototype.defer = Function.prototype.delay.curry(0.01);

Date.prototype.toJSON = function() {
	return '"' + this.getUTCFullYear() + '-' +
		(this.getUTCMonth() + 1).toPaddedString(2) + '-' +
		this.getUTCDate().toPaddedString(2) + 'T' +
		this.getUTCHours().toPaddedString(2) + ':' +
		this.getUTCMinutes().toPaddedString(2) + ':' +
		this.getUTCSeconds().toPaddedString(2) + 'Z"';
};

Object.extend(Date.prototype, {
	dateAdd: function(interval, n, cal) {
		var dt = new Date(this);
		if (!interval || !n) return;
		var s = 1, m = 1, h = 1, dd = 1, w = 1, i = interval;
		
		//	if the user wants workdays, but didn't provide a calendar, create a generic calendar
		//if (i == "workday" || i == "w") {
			if (!cal) {
				cal = new Object();
				cal.isHoliday = function () { return false; };
			}
		//}
		
		//	if the calendar doesn't have an isHoliday() method, create a generic one
		if (typeof cal.isHoliday == "undefined") {
			cal.isHoliday = function () { return false; };
		}
		
		if (i == "month" || i == "m" || i == "quarter" || i == "q" || i == "year" || i == "y"){
			dt = new Date(dt);
			if (i == "month" || i == "m") dt.setMonth(dt.getMonth() + n);
			if (i == "quarter" || i == "q") dt.setMonth(dt.getMonth() + (n * 3));
			if (i == "year" || i == "y") dt.setFullYear(dt.getFullYear() + n);		
		}
		else if (i == "second" || i == "s" || i == "minute" || i == "n" || i == "hour" || i == "h" || i == "day" || i == "d" || i == "workday" || i == "w" || i == "week" || i == "ww") {
			dt = Date.parse(dt);
			if (isNaN(dt)) return;
			if (i == "second" || i == "s") s = n;
			if (i == "minute" || i == "n") {s = 60; m = n}
			if (i == "hour" || i == "h") {s = 60; m = 60; h = n};
			if (i == "day" || i == "d") {s = 60; m = 60; h = 24; dd = n};
			if (i == "week" || i == "ww") {s = 60; m = 60; h = 24; dd = 7; w = n};
			dt += (((((1000 * s) * m) * h) * dd) * w);
			dt = new Date(dt);
			if (i == "workday" || i == "w") {
				s = 60;
				m = 60;
				h = 24;
				if (n > 0) {
					dtTemp = new Date(dt);
					for (var v = 1; v <= n; v++) {
						dtTemp.setDate(dtTemp.getDate() + 1);
						if (dtTemp.getDay() % 6 == 0 || cal.isHoliday(dtTemp)) v--;
					}
				}
				else {
					dtTemp = new Date(dt);
					for (var v = 0; v >= n; v--) {
						if (dtTemp.getDay() % 6 == 0 || cal.isHoliday(dtTemp)) v++;
						dtTemp.setDate(dtTemp.getDate() - 1);
					}
				}
				dt = new Date(dtTemp);
			}
		}
		return dt;
	},
	
	dateDiff: function(interval, dt2, firstdow, cal){
		//	return the number of intervals crossed
		//	11:01 to 11:59 is 0 hours
		//	10:59 to 11:01 is 1 hour
		//	When computing for days, if hours are used, the utc offset for each 
		//			date must be used, otherwise 1/1/2009 to 7/1/2009 returns one 
		//			hour short because we set our clocks an hour ahead in the 
		//			spring.  This may equate to losing a day, week, or month in 
		//			the calculation.
		
		var dt1 = new Date(this);
		if (!interval || !dt1 || !dt2) return;
		var v, s = 1, m = 1, h = 1, dd = 1, i = interval;
		
		//	if the user wants workdays, but didn't provide a calendar, create a generic calendar
		//if (i == "workday" || i == "w") {
			if (!cal) {
				cal = new Object();
				cal.isHoliday = function () { return false; };
			}
		//}
		
		//	if the calendar doesn't have an isHoliday() method, create a generic one
		if (typeof cal.isHoliday == "undefined") {
			cal.isHoliday = function () { return false; };
		}
		
		//	default firstdow to Sunday
		if (!firstdow || isNaN(firstdow)) {
			firstdow = 0;
		}
		firstdow *= 1.0;
		
		//Console.Writeln("ok");
		//if(i == "month" || i == "m" || i == "quarter" || i == "q" || i == "year" || i == "y"){
			dt1 = new Date(dt1);
			dt2 = new Date(dt2);
			years = dt2.getFullYear() - dt1.getFullYear();
			switch (i) {
				case "year" :
				case "y" :
					v = years;
					break;
				case "quarter" :
				case "q" :
					v = Math.floor(dt2.getMonth() / 3) - Math.floor(dt1.getMonth() / 3);
					if (years != 0) v += (years * 4);
					break;
				case "month" :
				case "m" :
					v = (dt2.getMonth() + 1) - (dt1.getMonth() + 1);
					if (years != 0) v += (years * 12);
					break;
				case "week" :
				case "w" :
					dt1.setDate(dt1.getDate() - (dt1.getDay() - firstdow) - (dt1.getDay() < firstdow ? 7 : 0));
					dt2.setDate(dt2.getDate() - (dt2.getDay() - firstdow) - (dt2.getDay() < firstdow ? 7 : 0));
					v = dt1.dateDiff("d", dt2) / 7;
					break;
				case "workday" :
				case "wd" :
					v = 0;
					if (dt1 == dt2) {
						v++;
					}
					else if (dt1 < dt2) {
						for (var dtTemp = new Date(dt1); dtTemp <= dt2; ) {
							//	don't count weekends or holidays
							//	Since sunday.getDay() = 0 and saturday.getDay() = 6, date.getDay() % 6 = 0 for both
							if (dtTemp.getDay() % 6 != 0 && !cal.isHoliday(dtTemp)) v++;
							dtTemp.setDate(dtTemp.getDate() + 1);
						}
					}
					else {
						for (var dtTemp = new Date(dt2); dtTemp <= dt1; ) {
							if (dtTemp.getDay() % 6 != 0 && !cal.isHoliday(dtTemp)) v--;
							dtTemp.setDate(dtTemp.getDate() + 1);
						}
					}
					v--;
					break;
				case "day" :
				case "d" :
					//	adjust for a difference caused by daylight savings
					dt2.setHours(dt2.getHours() + ((dt1.getTimezoneOffset() - dt2.getTimezoneOffset()) / 60));
					dt1 = Date.parse(dt1);
					dt2 = Date.parse(dt2);
					//	truncate the date value to the previous hour break
					dt1 -= (dt1 % (1000 * 60 * 60));
					dt2 -= (dt2 % (1000 * 60 * 60));
					v = (dt2 - dt1) / (1000 * 60 * 60 * 24);
					break;
				case "hour" :
				case "h" :
					dt1 = Date.parse(dt1);
					dt2 = Date.parse(dt2);
					//	truncate the date value to the previous hour break
					dt1 -= (dt1 % (1000 * 60 * 60));
					dt2 -= (dt2 % (1000 * 60 * 60));
					v = (dt2 - dt1) / (1000 * 60 * 60);
					break;
				case "minute" :
				case "n" :
					dt1 = Date.parse(dt1);
					dt2 = Date.parse(dt2);
					//	truncate the date value to the previous minute break
					dt1 -= (dt1 % (1000 * 60));
					dt2 -= (dt2 % (1000 * 60));
					v = (dt2 - dt1) / (1000 * 60);
					break;
				case "second" :
				case "s" :
					dt1 = Date.parse(dt1);
					dt2 = Date.parse(dt2);
					//	truncate the date value to the previous second break
					dt1 -= (dt1 % 1000);
					dt2 -= (dt2 % 1000);
					v = (dt2 - dt1) / 1000;
					break;
			}
		//}
		return v;
	},
	
	format: function (DateFormat) {
	/*
		Object Method:	Date.format()
		
		Type:		JavaScript 
		Purpose:	Provides a method for formatting dates according to 
					a specific date format pattern string.
		Inputs:		object			Required	a Date object
					NumberFormat	Required	the string pattern defining how to format the number
												eg. "mm/dd/yyyy"
		
		Returns:	A string containing the formatted date.
		
		Revision History
		Date		Developer	Description
		
	*/
		if (!this.isDate()) return "Invalid date";
		//if (!this.valueOf())
		//	return " ";
		if (DateFormat == undefined)
			return(this.toLocaleDateString());
		
		var d = this;
		
		DateFormat = DateFormat.toLowerCase();
		
		var s = DateFormat.replace(/(yyyy|yy|mmmm|mmm|mm|m|dddd|ddd|dd|d|hh|h|nn|n|ss|s|a\/p|am\/pm)/gi,
			function($1) {
				switch ($1.toLowerCase()) {
					case "am":  
					case "pm":  return $1;
					case "yyyy": return d.getFullYear();
					case "yy": return d.getFullYear().toString().substr(d.getFullYear().toString().length - 2);
					case "mmmm": return Application.MonthName[d.getMonth()];
					case "mmm":  return Application.MonthName[d.getMonth()].substr(0, 3);
					case "mm":   return (d.getMonth() + 1).toPaddedString(2);
					case "m":   return (d.getMonth() + 1);
					case "dddd": return Application.WeekdayName[d.getDay()];
					case "ddd":  return Application.WeekdayName[d.getDay()].substr(0, 3);
					case "dd":   return d.getDate().toPaddedString(2);
					case "d":   return d.getDate();
					case "hh":   
						h = d.getHours() % 24;
						if (DateFormat.search(/am\/pm/i) != -1 || DateFormat.search(/a\/p/i) != -1) {
							h %= 12;
							h = (h == 0) ? 12 : h;
						}
						return h.toPaddedString(2);
					case "h":   
						h = d.getHours() % 24;
						if (DateFormat.search(/am\/pm/i) != -1 || DateFormat.search(/a\/p/i) != -1) {
							h %= 12;
							h = (h == 0) ? 12 : h;
						}
						return h;
					case "nn":   return d.getMinutes().toPaddedString(2);
					case "n":   return d.getMinutes();
					case "ss":   return d.getSeconds().toPaddedString(2);
					case "s":   return d.getSeconds();
					case "a/p":  return $1 == $1.toLowerCase() ? (d.getHours() < 12 ? "a" : "p") : (d.getHours() < 12 ? "a" : "p").toUpperCase();
					case "am/pm":  return $1 == $1.toLowerCase() ? (d.getHours() < 12 ? "am" : "pm") : (d.getHours() < 12 ? "am" : "pm").toUpperCase();
					default:  return $1;
				}
			}
		);
		
		return s;
	},
	
	getBienniumShort: function () {
	/*
		Object Method:	Date.getBienniumShort()
		
		Type:		JavaScript 
		Purpose:	Return the biennium in the abbreviated format (07-09).
		Inputs:		object			Required	a Date object
		
		Returns:	A string containing the short-format biennium.
		
		Revision History
		Date		Developer	Description
		2/3/2009	Doug Pulse	Original
	*/
		var strTemp = "";
		if (this.getFullYear() % 2 == 0) {
			//    even year
			//    get the 2-digit year before and after
			strTemp = (this.getFullYear() - 1).toString().substr(2) + "-" + (this.getFullYear() + 1).toString().substr(2);
		}
		else {
			//    odd year
			if (this.getMonth() < 6) {
				//    last 6 months of the biennium
				strTemp = (this.getFullYear() - 2).toString().substr(2) + "-" + this.getFullYear().toString().substr(2);
			}
			else {
				//    first 6 months of the biennium
				strTemp = this.getFullYear().toString().substr(2) + "-" + (this.getFullYear() + 2).toString().substr(2);
			}
		}
		return strTemp;
	},
	
	getBiennium: function () {
	/*
		Object Method:	Date.getBiennium()
		
		Type:		JavaScript 
		Purpose:	Return the biennium (2007-2009).
		Inputs:		object			Required	a Date object
		
		Returns:	A string containing the biennium.
		
		Revision History
		Date		Developer	Description
		2/3/2009	Doug Pulse	Original
	*/
		var strTemp = "";
		if (this.getFullYear() % 2 == 0) {
			//    even year
			//    get the 2-digit year before and after
			strTemp = (this.getFullYear() - 1).toString() + "-" + (this.getFullYear() + 1).toString();
		}
		else {
			//    odd year
			if (this.getMonth() < 6) {
				//    last 6 months of the biennium
				strTemp = (this.getFullYear() - 2).toString() + "-" + this.getFullYear().toString();
			}
			else {
				//    first 6 months of the biennium
				strTemp = this.getFullYear().toString() + "-" + (this.getFullYear() + 2).toString();
			}
		}
		return strTemp;
	},
	
	getFiscalYear: function () {
	/*
		Object Method:	Date.getFiscalYear()
		
		Type:		JavaScript 
		Purpose:	Return the fiscal year (2009).
		Inputs:		object			Required	a Date object
		
		Returns:	A string containing the fiscal year.
		
		Revision History
		Date		Developer	Description
		8/4/2009	Doug Pulse	Original
	*/
		var intTemp = "";
		//    odd year
		if (this.getMonth() < 6) {
			//    last 6 months of the fiscal year
			intTemp = this.getFullYear();
		}
		else {
			//    first 6 months of the fiscal year
			intTemp = this.getFullYear() + 1;
		}
		return intTemp;
	},
	
	getBienniumEndYear: function () {
	/*
		Object Method:	Date.getBienniumEndYear()
		
		Type:		JavaScript 
		Purpose:	Return the biennium end year ("2007-2009" = 2009) as an integer.
		Inputs:		object			Required	a Date object
		
		Returns:	A integer containing the biennium end year.
		
		Revision History
		Date		Developer	Description
		2/3/2009	Doug Pulse	Original
	*/
		var intTemp = "";
		if (this.getFullYear() % 2 == 0) {
			//    even year
			//    get the 4-digit year after
			intTemp = this.getFullYear() + 1;
		}
		else {
			//    odd year
			if (this.getMonth() < 6) {
				//    last 6 months of the biennium
				intTemp = this.getFullYear();
			}
			else {
				//    first 6 months of the biennium
				intTemp = this.getFullYear() + 2;
			}
		}
		return intTemp;
	},
	
	getBienniumBeginYear: function () {
	/*
		Object Method:	Date.getBienniumBeginYear()
		
		Type:		JavaScript 
		Purpose:	Return the biennium begin year ("2007-2009" = 2007) as an integer.
		Inputs:		object			Required	a Date object
		
		Returns:	A integer containing the biennium begin year.
		
		Revision History
		Date		Developer	Description
		9/23/2010	Doug Pulse	Original
	*/
		return this.getBienniumEndYear() - 2;
	},
	
	getFederalFiscalYear: function () {
	/*
		Object Method:	Date.getFederalFiscalYear()
		
		Type:		JavaScript 
		Purpose:	Return the federal fiscal year as an integer.
		Inputs:		object			Required	a Date object
		
		Returns:	A integer containing the federal fiscal year.
		
		Revision History
		Date		Developer	Description
		2/3/2009	Doug Pulse	Original
	*/
		var intTemp = "";
		if (this.getMonth() < 9) {
			//    last 9 months of the fed fiscal year
			intTemp = this.getFullYear();
		}
		else {
			//    first 3 months of the fed fiscal year
			intTemp = this.getFullYear() + 1;
		}
		return intTemp;
	},
	
	getLocalFromUTC: function () {
	/*
		Object Method:	Date.getLocalFromUTC()
		
		Type:		JavaScript 
		Purpose:	Convert the UTC time to the local time.
		Inputs:		object			Required	a Date object
		
		Returns:	A date object
		
		Revision History
		Date		Developer	Description
		7/16/2010	Doug Pulse	Original
	*/
		var d = new Date(this);
		d.setHours(d.getHours() - (d.getTimezoneOffset() / 60));
		
		return d;
	},
	
	isDate: function () {
		return Object.isDate(this);
	}
});

RegExp.prototype.match = RegExp.prototype.test;

RegExp.escape = function(str) {
	return String(str).replace(/([.*+?^=!:${}()|[\]\/\\])/g, '\\$1');
};

/*--------------------------------------------------------------------------*/

Object.extend(String, {
	interpret: function(value) {
		return value == null ? '' : String(value);
	},
	specialChar: {
		'\b': '\\b',
		'\t': '\\t',
		'\n': '\\n',
		'\f': '\\f',
		'\r': '\\r',
		'\\': '\\\\'
	}
});

Object.extend(String.prototype, {
	gsub: function(pattern, replacement) {
		var result = '', source = this, match;
		replacement = arguments.callee.prepareReplacement(replacement);
		
		while (source.length > 0) {
			if (match = source.match(pattern)) {
				result += source.slice(0, match.index);
				result += String.interpret(replacement(match));
				source  = source.slice(match.index + match[0].length);
			} else {
				result += source, source = '';
			}
		}
		return result;
	},
	
	sub: function(pattern, replacement, count) {
		replacement = this.gsub.prepareReplacement(replacement);
		count = Object.isUndefined(count) ? 1 : count;
		
		return this.gsub(pattern, function(match) {
			if (--count < 0) return match[0];
			return replacement(match);
		});
	},
	
	scan: function(pattern, iterator) {
		this.gsub(pattern, iterator);
		return String(this);
	},
	
	truncate: function(length, truncation) {
		length = length || 30;
		truncation = Object.isUndefined(truncation) ? '...' : truncation;
		return this.length > length ?
		this.slice(0, length - truncation.length) + truncation : String(this);
	},
	
	strip: function() {
		return this.replace(/^\s+/, '').replace(/\s+$/, '');
	},
	
	stripTags: function() {
		return this.replace(/<\/?[^>]+>/gi, '');
	},
	
	evalScripts: function() {
		return this.extractScripts().map(function(script) { return eval(script) });
	},
	
	escapeHTML: function() {
		var self = arguments.callee;
		self.text.data = this;
		return self.div.innerHTML;
	},
	
	unescapeHTML: function() {
		var div = new Element('div');
		div.innerHTML = this.stripTags();
		return div.childNodes[0] ? (div.childNodes.length > 1 ?
		$A(div.childNodes).inject('', function(memo, node) { return memo+node.nodeValue }) :
		div.childNodes[0].nodeValue) : '';
	},
	
	toQueryParams: function(separator) {
		var match = this.strip().match(/([^?#]*)(#.*)?$/);
		if (!match) return { };
		
		return match[1].split(separator || '&').inject({ }, function(hash, pair) {
			if ((pair = pair.split('='))[0]) {
				var key = decodeURIComponent(pair.shift());
				var value = pair.length > 1 ? pair.join('=') : pair[0];
				if (value != undefined) value = decodeURIComponent(value);
				
				if (key in hash) {
					if (!Object.isArray(hash[key])) hash[key] = [hash[key]];
					hash[key].push(value);
				}
				else hash[key] = value;
			}
			return hash;
		});
	},
	
	toArray: function() {
		return this.split('');
	},
	
	toByteArray: function () {
		var a = new Array();
		for (var i = 0; i < this.length; i++) {
			a.push(this.charCodeAt(i));
		}
		return a;
	},
	
	succ: function() {
		return this.slice(0, this.length - 1) +
		String.fromCharCode(this.charCodeAt(this.length - 1) + 1);
	},
	
	times: function(count) {
		return count < 1 ? '' : new Array(count + 1).join(this);
	},
	
	camelize: function() {
		var parts = this.split('-'), len = parts.length;
		if (len == 1) return parts[0];
		
		var camelized = this.charAt(0) == '-'
			? parts[0].charAt(0).toUpperCase() + parts[0].substring(1)
			: parts[0];
			
		for (var i = 1; i < len; i++)
			camelized += parts[i].charAt(0).toUpperCase() + parts[i].substring(1);
		
		return camelized;
	},
	
	capitalize: function() {
		return this.charAt(0).toUpperCase() + this.substring(1).toLowerCase();
	},
	
	underscore: function() {
		return this.gsub(/::/, '/').gsub(/([A-Z]+)([A-Z][a-z])/,'#{1}_#{2}').gsub(/([a-z\d])([A-Z])/,'#{1}_#{2}').gsub(/-/,'_').toLowerCase();
	},
	
	dasherize: function() {
		return this.gsub(/_/,'-');
	},
	
	inspect: function(useDoubleQuotes) {
		var escapedString = this.gsub(/[\x00-\x1f\\]/, function(match) {
			var character = String.specialChar[match[0]];
			return character ? character : '\\u00' + match[0].charCodeAt().toPaddedString(2, 16);
		});
		if (useDoubleQuotes) return '"' + escapedString.replace(/"/g, '\\"') + '"';
		return "'" + escapedString.replace(/'/g, '\\\'') + "'";
	},
	
	toJSON: function() {
		return this.inspect(true);
	},
	
	unfilterJSON: function(filter) {
		return this.sub(filter || Hyperion.JSONFilter, '#{1}');
	},
	
	isJSON: function() {
		var str = this;
		if (str.blank()) return false;
		str = this.replace(/\\./g, '@').replace(/"[^"\\\n\r]*"/g, '');
		return (/^[,:{}\[\]0-9.\-+Eaeflnr-u \n\r\t]*$/).test(str);
	},
	
	evalJSON: function(sanitize) {
		var json = this.unfilterJSON();
		try {
			if (!sanitize || json.isJSON()) return eval('(' + json + ')');
		} catch (e) { }
		throw new SyntaxError('Badly formed JSON string: ' + this.inspect());
	},
	
	include: function(pattern) {
		return this.indexOf(pattern) > -1;
	},
	
	startsWith: function(pattern) {
		return this.indexOf(pattern) === 0;
	},
	
	endsWith: function(pattern) {
		var d = this.length - pattern.length;
		return d >= 0 && this.lastIndexOf(pattern) === d;
	},
	
	empty: function() {
		return this == '';
	},
	
	blank: function() {
		return /^\s*$/.test(this);
	},
	
	interpolate: function(object, pattern) {
		return new Template(this, pattern).evaluate(object);
	},
	
	column: function (width, just) {
		var s = this.toString();
		if (typeof width == "number") {
			if (typeof just == "undefined") {
				just = 0;
			}
			if (typeof just == "number") {
				if ((new Array(0, 1)).contains(just) && width > 0) {
					if (just == 0) {
						s += s.space(width);
						s = s.substr(0, width);
					}
					if (just == 1) {
						s = s.space(width) + s;
						s = s.substr(s.length - width);
					}
				}
			}
		}
		return s;
	},
	
	isNumeric:  function () {
		if (isNaN(parseFloat(this))) {
			return false;
		}
		return true;
	},
	
	isInteger: function() {
		return Object.isInteger(this);
	},
	
	reverse: function() {
	/*
		Object Method: String.reverse()
		
		Type: 		JavaScript 
		Purpose:	Provides a method for a String object to returns a string 
					in which the character order of the string is reversed.
		
		Inputs: 	(none)
		
		Returns:	string
		Modification History:
		Date		Developer	Description
		1/16/2008	Doug Pulse		Original
	*/
	  for( var oStr = "", x = this.length - 1, oTmp; oTmp = this.charAt(x); x-- ) {
	    oStr += oTmp;
	  }
	  return oStr;
	},
	
	lTrim: function() {
	/*
		Object Method: String.lTrim()
		
		Type:		JavaScript 
		Purpose:	Provides a method for the String object to 
					return a copy of a string without leading spaces.
		
		Inputs:		(none)
		
		Returns:	string
		Modification History:
		Date		Developer		Description
		1/16/2008	Doug Pulse		Original
	*/
		var whitespace = new String(" \t\n\r");
		
		var s = new String(this);
		
		if (whitespace.indexOf(s.charAt(0)) != -1) {
			// We have a string with leading blank(s)...
			
			var j = 0, i = s.length;
			
			// Iterate from the far left of string until we
			// don't have any more whitespace...
			while (j < i && whitespace.indexOf(s.charAt(j)) != -1)
				j++;
			
			// Get the substring from the first non-whitespace
			// character to the end of the string...
			s = s.substring(j, i);
		}
		
		return s;
	},
	
	rTrim: function() {
	/*
		Object Method: String.rTrim()
		
		Type:		JavaScript 
		Purpose:	Provides a method for the String object to 
					return a copy of a string without trailing spaces.
		
		Inputs: 	(none)
		
		Returns:	string
		Modification History:
		Date		Developer		Description
		1/16/2008	Doug Pulse		Original
	*/
		// We don't want to trim JUST spaces, but also tabs,
		// line feeds, etc.  Add anything else you want to
		// "trim" here in Whitespace
		var whitespace = new String(" \t\n\r");
		
		var s = new String(this);
		
		if (whitespace.indexOf(s.charAt(s.length - 1)) != -1) {
			// We have a string with trailing blank(s)...
			
			var i = s.length - 1;       // Get length of string
			
			// Iterate from the far right of string until we
			// don't have any more whitespace...
			while (i >= 0 && whitespace.indexOf(s.charAt(i)) != -1)
				i--;
			
			// Get the substring from the front of the string to
			// where the last non-whitespace character is...
			s = s.substring(0, i + 1);
		}
		
		return s;
	},
	
	space: function (width) {
		var s = "";
		if (typeof width == "number") {
			if (width > 0) {
				for (var i = 0; i < width; i++) {
					s += " ";
				}
			}
		}
		return s;
	},
	
	trim: function() {
	/*
		Object Method: String.trim()
		
		Type:		JavaScript 
		Purpose:	Provides a method for the String object to 
					return a copy of a string without leading or trailing spaces.
		
		Inputs:		(none)
		
		Returns:	string
		Modification History:
		Date		Developer		Description
		1/16/2008	Doug Pulse		Original
	*/
	 return this.lTrim().rTrim();
	},
	
	left: function(n) {
	/*
		Object Method: String.left()
		
		Type:		JavaScript 
		Purpose:	Provides a method for the String object to return a specified 
					number of characters from the left side of a string.
		
		Inputs: 	length	Numeric expression indicating how many characters to 
							return. If 0, a zero-length string("") is returned. If 
							greater than or equal to the number of characters in 
							string, the entire string is returned.
		
		Returns:	string
		
		Modification History:
		Date		Developer		Description
		1/16/2008	Doug Pulse		Original
	*/
		return this.substr(0, n);
	},
	
	right: function(n) {
	/*
		Object Method: String.right()
		
		Type:		JavaScript 
		Purpose:	Provides a method for the String object to return a specified 
					number of characters from the right side of a string.
		
		Inputs: 	length	Numeric expression indicating how many characters to 
							return. If 0, a zero-length string("") is returned. If 
							greater than or equal to the number of characters in 
							string, the entire string is returned.
		
		Returns:	string
		
		Modification History:
		Date		Developer		Description
		1/16/2008	Doug Pulse		Original
	*/
		return this.substr(this.length - n);
	},
	
	inStr: function(intStart, strSearchFor, blnCompare) {
	/*
		Name:		inStr
		Purpose:	provide a String method with the same functionality as the VBScript InStr function.
		Returns:	the position of the first occurrence of one string within another
		Inputs:		intStart		Numeric expression that sets the starting position for each search.
					strSearchFor	String expression searched for.
					blnCompare		0 = binary comparison, 1 = text comparison
		
		Notes:	vbscript instr function is 1-based
				this javascript instr function is 0-based
		
		Revision History
		Date		Developer		Description
		1/16/2008	D. Pulse		inagural
	*/
	        var intLoc = -1;
	        var strSearch = new String(this);
	        if (arguments[2] == undefined) 
	                blnCompare = 0;
	        if (blnCompare == 1) {
	                strSearch = strSearch.toLowerCase();
	                strSearchFor = strSearchFor.toLowerCase();
	        }
	        
	        intLoc = strSearch.substr(intStart).indexOf(strSearchFor);
	        
	        if (intLoc > -1) {
	                intLoc += intStart;
	        }
	        return(intLoc);
	},
	
	toPascalCase: function () {
		var arr = this.toProperCase().replace(/[ _\-\t]+/gi, " ").split(" ");
		
		for (var i = 0; i < arr.length; i++) {
			arr[i] = arr[i].substr(0, 1).toUpperCase() + arr[i].substr(1);
		}
		
		return arr.join("");
	},
	
	toCamelCase: function () {
		var arr = this.toProperCase().replace(/[ _\-\t]+/gi, " ").split(" ");
		
		arr[0] = arr[0].toLowerCase();
		
		for (var i = 1; i < arr.length; i++) {
			arr[i] = arr[i].substr(0, 1).toUpperCase() + arr[i].substr(1);
		}
		
		return arr.join("");
	},
	
	toUnderscoreCase: function () {
		return this.toProperCase().replace(/[ _\-\t]+/gi, "_");
	},
	
	toDashCase: function () {
		return this.toProperCase().replace(/[ _\-\t]+/gi, "-");
	},
	
	toProperCase: function () {
		return this.replace(/[_\- ]/g, " ").replace(/([A-Z]+)([A-Z])([a-z0-9,]+)/g, "$1 $2$3").replace(/([^A-Z ])([A-Z])/g, "$1 $2");
	},
	
	toSentenceCase: function () {
		var arr = this.toProperCase().replace(/[ _\-\t]+/gi, " ").split(" ");
		
		arr[0] = arr[0].substr(0, 1).toUpperCase() + arr[0].substr(1);
		
		for (var i = 1; i < arr.length; i++) {
			arr[i] = arr[i].toLowerCase();
		}
		
		return arr.join(" ");
	},
	
	toIRColumnName: function () {
		return this.toUnderscoreCase().replace(/[\(\)\-\%]/g, "_");
	}
});

String.prototype.gsub.prepareReplacement = function(replacement) {
	if (Object.isFunction(replacement)) return replacement;
	var template = new Template(replacement);
	return function(match) { return template.evaluate(match) };
};

String.prototype.parseQuery = String.prototype.toQueryParams;


Application.Template = Class.create({
	initialize: function(template, pattern) {
		this.template = template.toString();
		this.pattern = pattern || Template.Pattern;
	},
	
	evaluate: function(object) {
		if (Object.isFunction(object.toTemplateReplacements))
		object = object.toTemplateReplacements();
		
		return this.template.gsub(this.pattern, function(match) {
			if (object == null) return '';
			
			var before = match[1] || '';
			if (before == '\\') return match[2];
			
			var ctx = object, expr = match[3];
			var pattern = /^([^.[]+|\[((?:.*?[^\\])?)\])(\.|\[|$)/;
			match = pattern.exec(expr);
			if (match == null) return before;
			
			while (match != null) {
				var comp = match[1].startsWith('[') ? match[2].gsub('\\\\]', ']') : match[1];
				ctx = ctx[comp];
				if (null == ctx || '' == match[3]) break;
				expr = expr.substring('[' == match[3] ? match[1].length : match[0].length);
				match = pattern.exec(expr);
			}
			
			return before + String.interpret(ctx);
		});
	}
});
Template.Pattern = /(^|.|\r|\n)(#\{(.*?)\})/;

Application.$break = { };

Application.Enumerable = {
	each: function(iterator, context) {
		var index = 0;
		iterator = iterator.bind(context);
		try {
			this._each(function(value) {
				iterator(value, index++);
			});
		} catch (e) {
			if (e != $break) throw e;
		}
		return this;
	},
	
	eachSlice: function(number, iterator, context) {
		iterator = iterator ? iterator.bind(context) : Hyperion.K;
		var index = -number, slices = [], array = this.toArray();
		while ((index += number) < array.length)
		slices.push(array.slice(index, index+number));
		return slices.collect(iterator, context);
	},
	
	all: function(iterator, context) {
		iterator = iterator ? iterator.bind(context) : Hyperion.K;
		var result = true;
		this.each(function(value, index) {
			result = result && !!iterator(value, index);
			if (!result) throw $break;
		});
		return result;
	},
	
	any: function(iterator, context) {
		iterator = iterator ? iterator.bind(context) : Hyperion.K;
		var result = false;
		this.each(function(value, index) {
			if (result = !!iterator(value, index))
			throw $break;
		});
		return result;
	},
	
	collect: function(iterator, context) {
		iterator = iterator ? iterator.bind(context) : Hyperion.K;
		var results = [];
		this.each(function(value, index) {
			results.push(iterator(value, index));
		});
		return results;
	},
	
	detect: function(iterator, context) {
		iterator = iterator.bind(context);
		var result;
		this.each(function(value, index) {
			if (iterator(value, index)) {
				result = value;
				throw $break;
			}
		});
		return result;
	},
	
	findAll: function(iterator, context) {
		iterator = iterator.bind(context);
		var results = [];
		this.each(function(value, index) {
			if (iterator(value, index))
			results.push(value);
		});
		return results;
	},
	
	grep: function(filter, iterator, context) {
		iterator = iterator ? iterator.bind(context) : Hyperion.K;
		var results = [];
		
		if (Object.isString(filter))
		filter = new RegExp(filter);
		
		this.each(function(value, index) {
			if (filter.match(value))
			results.push(iterator(value, index));
		});
		return results;
	},
	
	include: function(object) {
		if (Object.isFunction(this.indexOf))
		if (this.indexOf(object) != -1) return true;
		
		var found = false;
		this.each(function(value) {
			if (value == object) {
				found = true;
				throw $break;
			}
		});
		return found;
	},
	
	inGroupsOf: function(number, fillWith) {
		fillWith = Object.isUndefined(fillWith) ? null : fillWith;
		return this.eachSlice(number, function(slice) {
			while(slice.length < number) slice.push(fillWith);
			return slice;
		});
	},
	
	inject: function(memo, iterator, context) {
		iterator = iterator.bind(context);
		this.each(function(value, index) {
			memo = iterator(memo, value, index);
		});
		return memo;
	},
	
	invoke: function(method) {
		var args = $A(arguments).slice(1);
		return this.map(function(value) {
			return value[method].apply(value, args);
		});
	},
	
	max: function(iterator, context) {
		iterator = iterator ? iterator.bind(context) : Hyperion.K;
		var result;
		this.each(function(value, index) {
			value = iterator(value, index);
			if (result == null || value >= result)
			result = value;
		});
		return result;
	},
	
	min: function(iterator, context) {
		iterator = iterator ? iterator.bind(context) : Hyperion.K;
		var result;
		this.each(function(value, index) {
			value = iterator(value, index);
			if (result == null || value < result)
			result = value;
		});
		return result;
	},
	
	partition: function(iterator, context) {
		iterator = iterator ? iterator.bind(context) : Hyperion.K;
		var trues = [], falses = [];
		this.each(function(value, index) {
			(iterator(value, index) ?
			trues : falses).push(value);
		});
		return [trues, falses];
	},
	
	pluck: function(property) {
		var results = [];
		this.each(function(value) {
			results.push(value[property]);
		});
		return results;
	},
	
	reject: function(iterator, context) {
		iterator = iterator.bind(context);
		var results = [];
		this.each(function(value, index) {
			if (!iterator(value, index))
			results.push(value);
		});
		return results;
	},
	
	sortBy: function(iterator, context) {
		iterator = iterator.bind(context);
		return this.map(function(value, index) {
			return {value: value, criteria: iterator(value, index)};
		}).sort(function(left, right) {
			var a = left.criteria, b = right.criteria;
			return a < b ? -1 : a > b ? 1 : 0;
		}).pluck('value');
	},
	
	toArray: function() {
		return this.map();
	},
	
	zip: function() {
		var iterator = Hyperion.K, args = $A(arguments);
		if (Object.isFunction(args.last()))
		iterator = args.pop();

		var collections = [this].concat(args).map($A);
		return this.map(function(value, index) {
			return iterator(collections.pluck(index));
		});
	},
	
	size: function() {
		return this.toArray().length;
	},
	
	inspect: function() {
		return '#<Enumerable:' + this.toArray().inspect() + '>';
	}
};

Object.extend(Enumerable, {
	map:     Enumerable.collect,
	find:    Enumerable.detect,
	select:  Enumerable.findAll,
	filter:  Enumerable.findAll,
	member:  Enumerable.include,
	entries: Enumerable.toArray,
	every:   Enumerable.all,
	some:    Enumerable.any
});

Array.from = $A;

Object.extend(Array.prototype, Enumerable);

if (!Array.prototype._reverse) Array.prototype._reverse = Array.prototype.reverse;

Object.extend(Array.prototype, {
	_each: function(iterator) {
		for (var i = 0, length = this.length; i < length; i++)
		iterator(this[i]);
	},
	
	clear: function() {
		this.length = 0;
		return this;
	},
	
	first: function() {
		return this[0];
	},
	
	last: function() {
		return this[this.length - 1];
	},
	
	compact: function() {
		return this.select(function(value) {
			return value != null;
		});
	},
	
	flatten: function() {
		return this.inject([], function(array, value) {
			return array.concat(Object.isArray(value) ?
			value.flatten() : [value]);
		});
	},
	
	without: function() {
		var values = $A(arguments);
		return this.select(function(value) {
			return !values.include(value);
		});
	},
	
	reverse: function(inline) {
		return (inline !== false ? this : this.toArray())._reverse();
	},
	
	reduce: function() {
		return this.length > 1 ? this : this[0];
	},
	
	uniq: function(sorted) {
		return this.inject([], function(array, value, index) {
			if (0 == index || (sorted ? array.last() != value : !array.include(value)))
				array.push(value);
			return array;
		});
	},
	
	intersect: function(array) {
		return this.uniq().findAll(function(item) {
			return array.detect(function(value) { return item === value });
		});
	},
	
	clone: function() {
		return [].concat(this);
	},
	
	size: function() {
		return this.length;
	},
	
	inspect: function() {
		return '[' + this.map(Object.inspect).join(', ') + ']';
	},
	
	toJSON: function() {
		var results = [];
		this.each(function(object) {
			var value = Object.toJSON(object);
			if (!Object.isUndefined(value)) results.push(value);
		});
		return '[' + results.join(', ') + ']';
	},
	
	contains: function (ValueToFind) {
	/*
		Object Method:	Array.contains()
		
		Type:		JavaScript 
		Purpose:	Determine if an element with the given value exists in an array
		Inputs:		object			Required	an array object
					ValueToFind		Required	the value to find
		
		Returns:	A boolean value indicating whether an element with a value
					matching the search criteria exists.
		
		Revision History
		Date		Developer	Description
		7/31/2008	D. Pulse	original
	*/
		for (var a in this) {
			if (this[a] == ValueToFind) return true;
		}
		for (var i = 0; i < this.length; i++) {
			if (this[i] == ValueToFind) return true;
		}
		return false;
	},
	
	demote: function (index) {
	/*
		Object Method:	Array.demote()
		
		Type:		JavaScript 
		Purpose:	Provides a method for moving an element closer to the end of the array
		Inputs:		object			Required	an array object
					index			Required	index of the element to demote
		
		Returns:	(none)
		
		Revision History
		Date		Developer	Description
		6/9/2008	D. Pulse	original
	*/
		var vTemp;
		if (index < this.length - 1) {
			vTemp = this[index + 1];
			this[index + 1] = this[index];
			this[index] = vTemp;
		}
	},
	
	promote: function (index) {
	/*
		Object Method:	Array.promote()
		
		Type:		JavaScript 
		Purpose:	Provides a method for moving an element closer to the beginning of the array
		Inputs:		object			Required	an array object
					index			Required	index of the element to promote
		
		Returns:	(none)
		
		Revision History
		Date		Developer	Description
		6/9/2008	D. Pulse	original
	*/
		var vTemp;
		if (index > 0) {
			vTemp = this[index - 1];
			this[index - 1] = this[index];
			this[index] = vTemp;
		}
	},
	
	eq: function (arr) {
	/*
		Object Method:	Array.eq()
		
		Type:		JavaScript 
		Purpose:	Provides a method for comparing two arrays
		Inputs:		object			Required	an array object
					arr				Required	an array object
		
		Returns:	A boolean value indicating whether arr is an array, Array 
					and arr are of the same length, and Array and arr contain 
					the same values.
		
		Revision History
		Date		Developer	Description
		6/9/2008	D. Pulse	original
	*/
		if (arr instanceof Array) {
		//	Is arr and array?
			if (this.length == arr.length) {
			//	Are this array and arr of the same length
				for (var i = 0; i < this.length; i++) {
				//	Do all of the numerically-referenced values match?
					if (this[i] instanceof Array) {
						if (!this[i].eq(arr[i])) {
							return(false);
						}
					}
					else {
						if (this[i] != arr[i]) {
							return(false);
						}
					}
				}
				for (a in this) {
				//	Do all of the associatively-referenced values match?
					if (this[a] instanceof Array) {
						if (!this[a].eq(arr[a])) {
							return(false);
						}
					}
					else {
						if (this[a] != arr[a]) {
							return(false);
						}
					}
				}
				for (a in arr) {
				//	Do all of the associatively-referenced values match?
					if (this[a] instanceof Array) {
						if (!this[a].eq(arr[a])) {
							return(false);
						}
					}
					else {
						if (arr[a] != this[a]) {
							return(false);
						}
					}
				}
			}
			else {
				return(false);
			}
		}
		else {
			return(false);
		}
		return(true);
	},
	
//	indexOf: function (value, startIndex) {
//	/*
//		Object Method:	Array.indexOf()
//		
//		Type:		JavaScript 
//		Purpose:	Provides a method for locating the first element of an array
//					that contains a specific value.
//					(Same as Array.indicesOf[0].)
//		Inputs:		object			Required	an array object
//					value			Required	the value to match
//		
//		Returns:	The index of the first element that contains the specified value.
//		
//		Revision History
//		Date		Developer	Description
//		6/9/2008	D. Pulse	original
//	*/
//		var n = 0;
//		if (typeof startIndex == "number") {
//			if (Math.floor(startIndex) > -1 && Math.floor(startIndex) < this.length) {
//				n = Math.floor(startIndex);
//			}
//		}
//		for (var i = n; i < this.length; i++) {
//			if (this[i] == value) break;
//		}
//		return i == this.length ? -1 : i;
//	},
	
	indicesOf: function (value) {
	/*
		Object Method:	Array.indicesOf()
		
		Type:		JavaScript 
		Purpose:	Provides a method for locating all elements of an array
					that contain a specific value.
		Inputs:		object			Required	an array object
					value			Required	the value to match
		
		Returns:	An array containing the numeric indices of the elements
					that contain values matching the one specified.
		
		Revision History
		Date		Developer	Description
		6/9/2008	D. Pulse	original
	*/
		var a = new Array();
		for (var i = 0; i < this.length; i++) {
			if (this[i] == value) a.push(i);
		}
		return a;
	},
	
	remove: function (value, count) {
	/*
		Object Method:	Array.remove()
		
		Type:		JavaScript 
		Purpose:	Removes from the array elements whose values match the one specified.
		Inputs:		object			Required	an array object
					value			Required	the value to match
					count			Optional	The number of elements to remove.
												If count is omitted, all matching 
												elements are removed.
		
		Returns:	An array containing the numeric indices of the elements
					that contain values matching the one specified.
		
		Revision History
		Date		Developer	Description
		6/9/2008	D. Pulse	original
	*/
		var a = this.indicesOf(value);
		var n = a.length;
		if (typeof count == "number") {
			if (Math.floor(count) > 0 && Math.floor(count) < n) {
				n = Math.floor(count);
			}
		}
		a = a.slice(0, n);
		for (var i = a.length - 1; i > -1; i--) {
			this.splice(a[i], 1);
			n--;
			if (n < 1) break;
		}
		return n;
	}
});

Object.extend(Array.prototype, {
	sum: function () {
	/*
		Object Method:	Array.sum()
		
		Type:		JavaScript 
		Purpose:	Sum the values in the array.
		Inputs:		object		Required	an array object
		
		Returns:	The result of summing all of the array's values.
		
		Comments:	This method is fragile.  Passing arrays containing elements of differing types
					may cause undesirable results.
					We really should check to be sure that all of the values are of the same data type
					We can sum numbers or strings (concatenate), but not objects
		
		Revision History
		Date		Developer	Description
		5/11/2010	D. Pulse	original
	*/
		var _a;
		var _type;
		var _err = false;
		
		if (this.length > 0) {
			_type = typeof this[0];
			if (_type == "number" || _type == "string") {
				for (var i = 1; i < this.length; i++) {
					if (typeof this[i] != _type) {
						_err = true;
						break;
					}
				}
				if (!_err) {
					_a = _type == "string" ? "" : 0;
					
					for (var i = 0; i < this.length; i++) {
						_a += this[i];
					}
				}
			}
		}
		return _a;
	},
	
	and: function (arr) {
	/*
		Object Method:	Array.and()
		
		Type:		JavaScript 
		Purpose:	Find which elements are in both arrays.
		Inputs:		object		Required	an array object
					arr			optional	The Array object to compare to.
		
		Returns:	An array containing all of the values that are in arr and the original array.
		
		Revision History
		Date		Developer	Description
		7/9/2010	D. Pulse	original
	*/
		var arrOut = new Array();
		
		if (Array.prototype.isPrototypeOf(arr)) {
			for (var i = 0; i < this.length; i++) {
				if (arr.contains(this[i])) arrOut.push(this[i]);
			}
		}
		
		return arrOut.unique();
	},
	
	or: function (arr) {
	/*
		Object Method:	Array.or()
		
		Type:		JavaScript 
		Purpose:	Find which elements are in either array.
		Inputs:		object		Required	an array object
					arr			optional	The Array object to compare to.
		
		Returns:	An array containing all of the values that are in arr or the original array.
		
		Revision History
		Date		Developer	Description
		7/9/2010	D. Pulse	original
	*/
		var arrOut = this.clone().unique();
		
		if (Array.prototype.isPrototypeOf(arr)) {
			arrOut = this.concat(arr);
		}
		return arrOut.unique();
	},
	
	xor: function (arr) {
	/*
		Object Method:	Array.xor()
		
		Type:		JavaScript 
		Purpose:	Find which elements are in either array, but not both.
		Inputs:		object		Required	an array object
					arr			optional	The Array object to compare to.
		
		Returns:	An array containing all of the values that are in arr or the original array but not in both.
		
		Revision History
		Date		Developer	Description
		7/9/2010	D. Pulse	original
	*/
		var arrAnd = new Array();
		var arrOr = new Array();
		var arrOut = this.clone().unique();
		
		if (Array.prototype.isPrototypeOf(arr)) {
			arrOr = this.or(arr);
			arrAnd = this.and(arr);
			for (var i = 0; i < arrAnd.length; i++) {
				arrOr.remove(arrAnd[i]);
			}
			arrOut = arrOr.clone();
		}
		
		return arrOut;
	},
	
	mean: function () {
	/*
		Object Method:	Array.mean()
		
		Type:		JavaScript 
		Purpose:	Find the arithmetic mean of an array of numbers.
		Inputs:		object		Required	an array object
		
		Returns:	A numerical value representing the mean of the array of numbers.
		
		Revision History
		Date		Developer	Description
		9/23/2010	D. Pulse	original
	*/
		//	verify that all of the array's elements are of the same type
		for (var i = 1; i < this.length; i++) {
			if ((typeof this[i]) != (typeof this[i])) return;
		}
		
		var vOut;
		
		if (typeof this[0] == "number") {
			vOut = 1.0 * this.sum() / this.length;
		}
		if (Date.prototype.isPrototypeOf(this[0])) {
			var arrDV = new Array();
			for (var i = 0; i < this.length; i++) {
				arrDV.push(this[i].valueOf());
			}
			vOut = new Date(arrDV.mean());
		}
		//	other data types are not supported
		
		return vOut;
	}, //	Array.mean
	
	median: function () {
	/*
		Object Method:	Array.median()
		
		Type:		JavaScript 
		Purpose:	Find the median value of an array of numbers.
		Inputs:		object		Required	an array object
		
		Returns:	A numerical value representing the median of the array of numbers.
		
		Revision History
		Date		Developer	Description
		9/23/2010	D. Pulse	original
	*/
		//	verify that all of the array's elements are of the same type
		for (var i = 1; i < this.length; i++) {
			if ((typeof this[i]) != (typeof this[i])) return;
		}
		
		var vOut;
		var a = this.clone();
		a.sort();
		if (typeof this[0] == "number") {
			if (this.length % 2 == 0) {
				vOut = (a[(a.length / 2) - 1] + a[a.length / 2]) / 2;
			}
			else {
				vOut = a[(a.length - 1) / 2];
			}
		}
		if (Date.prototype.isPrototypeOf(this[0])) {
			var arrDV = new Array();
			for (var i = 0; i < this.length; i++) {
				arrDV.push(this[i].valueOf());
			}
			vOut = new Date(arrDV.median());
		}
		
		return vOut;
	}, //	Array.median
	
	mode: function () {
	/*
		Object Method:	Array.mode()
		
		Type:		JavaScript 
		Purpose:	Find the mode of an array of values.
		Inputs:		object		Required	an array object
		
		Returns:	An array containing the mode(s) of the array of values.
		
		Comment:	There are 0 or more modes in a set of values.
		
		Revision History
		Date		Developer	Description
		9/23/2010	D. Pulse	original
	*/
		//	verify that all of the array's elements are of the same type
		for (var i = 1; i < this.length; i++) {
			if ((typeof this[i]) != (typeof this[i])) return;
		}
		
		var a = new Array();
		var b = new Array();
		var c = new Array();
		
		//	How many of each value?
		for (var i = 0; i < this.length; i++) {
			if (a.contains(this[i])) {
				//	the value has been seen before
				//	increment its counter
				b[a.indexOf(this[i])]++;
			}
			else {
				//	this is a new value
				//	create a counter for it
				a.push(this[i]);
				b.push(1);
			}
		}
		
		//	find the most common value
		var arrOut = new Array();
		if (b.max() == 1) {
			
		}
		else {
			c = b.indicesOf(b.max());
			for (var i = 0; i < c.length; i++) {
				arrOut.push(a[c[i]]);
			}
		}
		
		return arrOut;
	}, //	Array.mode
	
	range: function () {
	/*
		Object Method:	Array.range()
		
		Type:		JavaScript 
		Purpose:	Find the arithmetic range of an array of numbers.
		Inputs:		object		Required	an array object
		
		Returns:	A numerical value representing the range of the array of numbers.
		
		Comment:	If the array contains dates, the return value is the number of days.
		
		Revision History
		Date		Developer	Description
		9/23/2010	D. Pulse	original
	*/
		//	verify that the array contains only numbers
		for (var i = 0; i < this.length; i++) {
			if (isNaN(this[i])) return;
		}
		
		return this.max() - this.min();
	},  //	Array.range
	
	stdDev: function (isSample) {
	/*
		Object Method:	Array.stdDev()
		
		Type:		JavaScript 
		Purpose:	Find the standard deviation of an array of numbers.
		Inputs:		object		Required	an array object
					isSample	Optional	A boolean value indicating whether the array contains sample data.
											true = sample
											false (or omitted) = population
		
		Returns:	A numerical value representing the standard deviation of the array of numbers.
		
		Revision History
		Date		Developer	Description
		4/25/2013	D. Pulse	original
	*/
		if (typeof isSample == "undefined") isSample = false;
		if (typeof isSample != "boolean") isSample = false;
		var a = isSample ? 1 : 0;
		
		//	verify that the array contains only numbers
		for (var i = 0; i < this.length; i++) {
			if (isNaN(this[i])) return -1;
			if (!Object.isNumber(this[i])) return -1;
		}
		
		var m = this.mean();
		
		var arrDelta = new Array();
		for (var i = 0; i < this.length; i++) {
			arrDelta.push((new Number(this[i].valueOf() - m)).pow(2));
		}
		
		return ( arrDelta.sum() / (this.length - a) ).pow(0.5);
	},  //	Array.stdDev
	
	skewness: function (useFisherPearson) {
	/*
		Object Method:	Array.skewness()
		
		Type:		JavaScript 
		Purpose:	Find the sample skewness of an array of numbers.
		Inputs:		object				Required	an array object
					useFisherPearson	Optional	A boolean value indicating whether to use the adjusted 
													Fisher-Pearson standardized moment coefficient, which is what is
													used by Excel and several statistical packages like Minitab, SAS 
													and SSPS.
													true = use adjusted Fisher-Pearson
													false (or omitted) = use the usual estimator
		
		Returns:	A numerical value representing the standard deviation of the array of numbers.
		Reference:	http://en.wikipedia.org/wiki/Skewness
		
		Revision History
		Date		Developer	Description
		4/25/2013	D. Pulse	original
	*/
		if (typeof useFisherPearson == "undefined") useFisherPearson = false;
		if (typeof useFisherPearson != "boolean") useFisherPearson = false;
		var m = this.mean();
		var n = this.length;
		
		if (useFisherPearson) {
			//	adjusted Fisher-Pearson
			var arrMS3 = [];
			var s = this.stdDev();
			for (var i = 0; i < this.length; i++) {
				arrMS3.push(((this[i] - m)/s).pow(3));
			}
			var ms3 = arrMS3.sum();
			var G = (n * ms3) / ((n - 1) * (n - 2));
		}
		else {
			//	usual
			var arrM3 = [];
			var arrM2 = [];
			for (var i = 0; i < this.length; i++) {
				arrM3.push((this[i] - m).pow(3));
				arrM3.push((this[i] - m).pow(2));
			}
			var m3 = arrM3.sum() / n;
			var m2 = (arrM2.sum() / n);
			var g1 = m3 / m2.pow(1.5);
			var G = ((n * (n - 1)).pow(0.5) / (n - 2)) * g1;
		}
		return G;
	}, //	Array.skewness
	
	sampleSize: function (cl, ci) {
	/*
		Object Method:	Array.sampleSize()
		
		Type:		JavaScript 
		Purpose:	Find the required sample size based on sample data.
		Inputs:		object	Required	an array object
					cl		Required	confidence level
										Z-values are hard coded so only cl = 0.95 or 0.99 is allowed.
					ci		Required	confidence interval
										Assuming results are between 0 and 1, ci must be between 0 and 0.5.
		
		Returns:	A numerical value representing the standard deviation of the array of numbers.
		Reference:	http://www.surveysystem.com/sscalc.htm#ssneeded
		
		Revision History
		Date		Developer	Description
		4/26/2013	D. Pulse	original
	*/
		var ss = new Number(0);
		if ([0.95, 0.99].contains(cl)) {
			if (typeof ci == "number") {
				if (ci > 0 && ci <= 0.5) {
					var zVal = cl == 0.95 ? 1.96 : 2.58;
					ss = (zVal * this.stdDev(true) / ci).pow(2);
				}
			}
		}
		
		return Math.round(ss);
	}	//	Array.sampleSize
});

// use native browser JS 1.6 implementation if available
if (Object.isFunction(Array.prototype.forEach))
	Array.prototype._each = Array.prototype.forEach;

if (!Array.prototype.indexOf) Array.prototype.indexOf = function(item, i) {
	i || (i = 0);
	var length = this.length;
	if (i < 0) i = length + i;
	for (; i < length; i++)
	if (this[i] === item) return i;
	return -1;
};

if (!Array.prototype.lastIndexOf) Array.prototype.lastIndexOf = function(item, i) {
	i = isNaN(i) ? this.length : (i < 0 ? this.length + i : i) + 1;
	var n = this.slice(0, i).reverse().indexOf(item);
	return (n < 0) ? n : i - n - 1;
};

Array.prototype.toArray = Array.prototype.clone;
//Array.prototype.unique = Array.prototype.uniq;
Array.prototype.unique = function(sorted) {
	var arr = this.uniq();
	if (sorted) arr = arr.sort();
	return arr;
}

Application.$w = function (string) {
	if (!Object.isString(string)) return [];
	string = string.strip();
	return string ? string.split(/\s+/) : [];
}

Object.extend(Number.prototype, {
	toColorPart: function() {
		return this.toPaddedString(2, 16);
	},
	
	succ: function() {
		return this + 1;
	},
	
	times: function(iterator) {
		$R(0, this, true).each(iterator);
		return this;
	},
	
	toPaddedString: function(length, radix) {
		var string = this.toString(radix || 10);
		return '0'.times(length - string.length) + string;
	},
	
	toJSON: function() {
		return isFinite(this) ? this.toString() : 'null';
	},
	
	format: function (NumDigitsAfterDecimal, IncludeLeadingDigit, UseParensForNegativeNumbers, GroupDigits) {
	/*
		Object Method: Number.format()
		
		Type:		JavaScript 
		Purpose:	Provides a method for the Number object to display a number in a specific format.
		Inputs:		object						Required	a Number object
					NumDigitsAfterDecimal		Optional	the number of decimal places to format the number to
					IncludeLeadingDigit			Optional	true / false - display a leading zero for
															numbers between -1 and 1
					UseParensForNegativeNumbers	Optional	true / false - use parenthesis around negative numbers
					GroupDigits					Optional	put commas as number separators.
		
		Returns:	A string containing the formatted number.
		Modification History:
		Date		Developer		Description
		1/16/2008	Doug Pulse		Original
	*/
		if (typeof NumDigitsAfterDecimal == "undefined") NumDigitsAfterDecimal = 0;
		if (typeof IncludeLeadingDigit == "undefined") IncludeLeadingDigit = true;
		if (typeof UseParensForNegativeNumbers == "undefined") UseParensForNegativeNumbers = false;
		if (typeof GroupDigits == "undefined") GroupDigits = false;
		
		if (isNaN(this.valueOf())) return "NaN";
		
		var tmpNum = this.valueOf();
		var iSign = this.valueOf() < 0 ? -1 : 1;		// Get sign of number
		
		// Adjust number so only the specified number of numbers after
		// the decimal point are shown.
		tmpNum *= Math.pow(10, NumDigitsAfterDecimal);
		tmpNum = Math.round(Math.abs(tmpNum));
		tmpNum /= Math.pow(10, NumDigitsAfterDecimal);
		tmpNum *= iSign;					// Readjust for sign
		
		// Create a string object to do our formatting on
		var tmpNumStr = new String(tmpNum);
		
		if (tmpNumStr.indexOf(".") == -1)
			tmpNumStr += ".";
		
		while ((tmpNumStr.length - tmpNumStr.lastIndexOf(".") - 1) != NumDigitsAfterDecimal) {
			tmpNumStr += "0";
		}
		
		// See if we need to strip out the leading zero or not.
		if (!IncludeLeadingDigit && this.valueOf() < 1 && this.valueOf() > -1 && this.valueOf() != 0)
			if (this.valueOf() > 0)
				tmpNumStr = tmpNumStr.substring(1, tmpNumStr.length);
			else
				tmpNumStr = "-" + tmpNumStr.substring(2, tmpNumStr.length);
			
		// See if we need to put in the commas
		if (GroupDigits && (this.valueOf() >= 1000 || this.valueOf() <= -1000)) {
			var iStart = tmpNumStr.indexOf(".");
			if (iStart < 0)
				iStart = tmpNumStr.length;

			iStart -= 3;
			while (iStart >= 1) {
				tmpNumStr = tmpNumStr.substring(0, iStart) + "," + tmpNumStr.substring(iStart, tmpNumStr.length);
				iStart -= 3;
			}
		}

		// See if we need to use parenthesis
		if (UseParensForNegativeNumbers && this.valueOf() < 0)
			tmpNumStr = "(" + tmpNumStr.substring(1, tmpNumStr.length) + ")";
		
		// Make sure the number doesn't end with a decimal point
		if (tmpNumStr.substr(tmpNumStr.length - 1, 1) == ".")
			tmpNumStr = tmpNumStr.substr(0, tmpNumStr.length - 1);
		
		return tmpNumStr;		// Return our formatted string!
	},
	
	currencyFormat: function (NumDigitsAfterDecimal, IncludeLeadingDigit, UseParensForNegativeNumbers, GroupDigits) {
	/*
		Object Method:	Number.currencyFormat()
		
		Type:		JavaScript 
		Purpose:	Provides a method for the Number object to display a number as currency.
		Inputs:		object						Required	a Number object
					NumDigitsAfterDecimal		Optional	the number of decimal places to format the number to
					IncludeLeadingDigit			Optional	true / false - display a leading zero for
															numbers between -1 and 1
					UseParensForNegativeNumbers	Optional	true / false - use parenthesis around negative numbers
					GroupDigits					Optional	put commas as number separators.
		
		Returns:	A string containing the number formatted as currency.
		
		Revision History
		Date		Developer	Description
		
	*/
		var tmpStr = new String(this.format(NumDigitsAfterDecimal, IncludeLeadingDigit, UseParensForNegativeNumbers, GroupDigits));
		
		if (tmpStr.indexOf("(") != -1 || tmpStr.indexOf("-") != -1) {
			// We know we have a negative number, so place '$' inside of '(' / after '-'
			if (tmpStr.charAt(0) == "(")
				tmpStr = "($"  + tmpStr.substring(1,tmpStr.length);
			else if (tmpStr.charAt(0) == "-")
				tmpStr = "-$" + tmpStr.substring(1,tmpStr.length);
				
			return tmpStr;
		}
		else
			return "$" + tmpStr;		// Return formatted string!
	},
	
	percentFormat: function (NumDigitsAfterDecimal, IncludeLeadingDigit, UseParensForNegativeNumbers, GroupDigits) {
	/*
		Object Method:	Number.percentFormat()
		
		Type:		JavaScript 
		Purpose:	Provides a method for the Number object to display a number as a percent in a specific format.
		Inputs:		object						Required	a Number object
					NumDigitsAfterDecimal		Optional	the number of decimal places to format the number to
					IncludeLeadingDigit			Optional	true / false - display a leading zero for
															numbers between -1 and 1
					UseParensForNegativeNumbers	Optional	true / false - use parenthesis around negative numbers
					GroupDigits					Optional	put commas as number separators.
		
		Returns:	A string containing the number formatted as a percent
		Modification History:
		Date		Developer		Description
		1/16/2008	Doug Pulse		Original
	*/
		var tmpStr = new String((this * 100).format(NumDigitsAfterDecimal, IncludeLeadingDigit, UseParensForNegativeNumbers, GroupDigits));
		
		if (tmpStr.indexOf(")") != -1) {
			// We know we have a negative number, so place '%' inside of ')'
			tmpStr = tmpStr.substring(0, tmpStr.length - 1) + "%)";
			return tmpStr;
		}
		else
			return tmpStr + "%";			// Return formatted string!
	},
	
	stringFormat: function (NumberFormat) {
	/*
		Object Method:	Number.stringFormat()
		
		Type:		JavaScript 
		Purpose:	Provides a method for the Number object to display 
					a number in a specific format.
		Inputs:		object			Required	a Number object
					NumberFormat	Required	the string pattern defining how to format the number
												eg. "$#,##0"
		
		Returns:	A string containing the formatted number.
		
		Revision History
		Date		Developer	Description
		
	*/
		var iLoc;
		
		//if (NumberFormat.inStr(0, ",") != -1) 
		//	GroupDigits = 1;
		
		var strTemp = NumberFormat.replace(/(\#|0|,|.|\$)/gi, "");
		if (strTemp.length > 0) 
			return("Invalid number format");
		
		iLoc = NumberFormat.inStr(0, ".")
		if (iLoc != -1) {
			NumDigitsAfterDecimal = NumberFormat.length - iLoc - 1;
		}
		else {
			NumDigitsAfterDecimal = 0;
		}
		
		if (NumberFormat.substr(0, iLoc).inStr(0, "0") != 0) {
			IncludeLeadingDigit = 1;
		}
		else {
			IncludeLeadingDigit = 1;
		}
		
		iLoc = NumberFormat.inStr(0, ",")
		if (iLoc != -1) {
			GroupDigits = 1;
		}
		else {
			GroupDigits = 0;
		}
		
		iLoc = NumberFormat.inStr(0, "(")
		if (iLoc != -1) {
			UseParensForNegativeNumbers = 1;
		}
		else {
			UseParensForNegativeNumbers = 0;
		}
		
		iLoc = NumberFormat.inStr(0, "$")
		if (iLoc != -1) {
			return(this.currencyFormat(NumDigitsAfterDecimal, IncludeLeadingDigit, UseParensForNegativeNumbers, GroupDigits));
		}
		else {
			return(this.format(NumDigitsAfterDecimal, IncludeLeadingDigit, UseParensForNegativeNumbers, GroupDigits));
		}
	}, 
	
	pow: function (n) {
		return Math.pow(this, n);
	}
});

Number.prototype.zf = Number.prototype.toPaddedString;

$w('abs round ceil floor').each(function(method){
	Number.prototype[method] = Math[method].methodize();
});


Number.prototype.round = function(places) {
	if (typeof places == "undefined") {
		return(Math.round(this));
	}
	else if (typeof places == "number") {
		return Math.round(this * Math.pow(10, places)) / Math.pow(10, places);
	}
	else {
		return;
	}
};


Application.$H = function (object) {
	return new Hash(object);
};

Application.Hash = Class.create(Enumerable, (function() {
	
	function toQueryPair(key, value) {
		if (Object.isUndefined(value)) return key;
		return key + '=' + encodeURIComponent(String.interpret(value));
	}
	
	return {
		initialize: function(object) {
			this._object = Object.isHash(object) ? object.toObject() : Object.clone(object);
		},
		
		_each: function(iterator) {
			for (var key in this._object) {
				var value = this._object[key], pair = [key, value];
				pair.key = key;
				pair.value = value;
				iterator(pair);
			}
		},
		
		set: function(key, value) {
			return this._object[key] = value;
		},
		
		get: function(key) {
			return this._object[key];
		},
		
		unset: function(key) {
			var value = this._object[key];
			delete this._object[key];
			return value;
		},
		
		toObject: function() {
			return Object.clone(this._object);
		},
		
		keys: function() {
			return this.pluck('key');
		},
		
		values: function() {
			return this.pluck('value');
		},
		
		index: function(value) {
			var match = this.detect(function(pair) {
				return pair.value === value;
			});
			return match && match.key;
		},
		
		merge: function(object) {
			return this.clone().update(object);
		},
		
		update: function(object) {
			return new Hash(object).inject(this, function(result, pair) {
				result.set(pair.key, pair.value);
				return result;
			});
		},
		
		toQueryString: function() {
			return this.map(function(pair) {
				var key = encodeURIComponent(pair.key), values = pair.value;
				
				if (values && typeof values == 'object') {
					if (Object.isArray(values))
					return values.map(toQueryPair.curry(key)).join('&');
				}
				return toQueryPair(key, values);
			}).join('&');
		},
		
		inspect: function() {
			return '#<Hash:{' + this.map(function(pair) {
			return pair.map(Object.inspect).join(': ');
			}).join(', ') + '}>';
		},
		
		toJSON: function() {
			return Object.toJSON(this.toObject());
		},
		
		clone: function() {
			return new Hash(this);
		},
		
		toString:  function() {
			return this.inspect();
		}
	}
})());

Hash.prototype.toTemplateReplacements = Hash.prototype.toObject;
Hash.from = $H;

//Application.Dictionary = Class.create({
//	initialize:  function () {
//		_h = new Application.Hash();
//	},
//	
//	
//});


Application.ObjectRange = Class.create(Enumerable, {
	initialize: function(start, end, exclusive) {
		this.start = start;
		this.end = end;
		this.exclusive = exclusive;
	},

	_each: function(iterator) {
		var value = this.start;
		while (this.include(value)) {
			iterator(value);
			value = value.succ();
		}
	},

	include: function(value) {
		if (value < this.start)
			return false;
		if (this.exclusive)
			return value < this.end;
		return value <= this.end;
	}
});

Application.$R = function(start, end, exclusive) {
	return new ObjectRange(start, end, exclusive);
};

//	function $(element) {
//		if (arguments.length > 1) {
//			for (var i = 0, elements = [], length = arguments.length; i < length; i++)
//				elements.push($(arguments[i]));
//			return elements;
//		}
//		if (Object.isString(element))
//			element = document.getElementById(element);
//		return Element.extend(element);
//	}


Object.extend(Application, {
	AppType: ["Unknown", "DesktopClient", "PlugInClient", "ThinClient", "Scheduler", "SmartViewClient", "Migration"][Application.Type],
	FSO:  new JOOLEObject("Scripting.FileSystemObject"),
	Net:  new JOOLEObject("WScript.Network"),
	WShell:  new JOOLEObject("WScript.Shell")
});


//Object.extend(Application, {
//	JSON: {
//		parse: function (sJSON) { return eval("(" + sJSON + ")"); },
//		stringify: function (vContent) {
//			if (vContent instanceof Object) {
//				var sOutput = "";
//				if (vContent.constructor === Array) {
//					for (var nId = 0; nId < vContent.length; sOutput += this.stringify(vContent[nId]) + ",", nId++);
//					return "[" + sOutput.substr(0, sOutput.length - 1) + "]";
//				}
//				if (vContent.toString !== Object.prototype.toString) { return "\"" + vContent.toString().replace(/"/g, "\\$&") + "\""; }
//				for (var sProp in vContent) { sOutput += "\"" + sProp.replace(/"/g, "\\$&") + "\":" + this.stringify(vContent[sProp]) + ","; }
//				return "{" + sOutput.substr(0, sOutput.length - 1) + "}";
//			}
//			return typeof vContent === "string" ? "\"" + vContent.replace(/"/g, "\\$&") + "\"" : String(vContent);
//		}
//	}
//});


//	var WinVerHash = new Application.Hash();
//	WinVerHash.set("1.04", "Windows 1.0");
//	WinVerHash.set("2.11", "Windows 2.0");
//	WinVerHash.set("3", "Windows 3.0");
//	WinVerHash.set("3.10.528", "Windows NT 3.1");
//	WinVerHash.set("3.11", "Windows for Workgroups 3.11");
//	WinVerHash.set("3.5.807", "Windows NT Workstation 3.5");
//	WinVerHash.set("3.51.1057", "Windows NT Workstation 3.51");
//	WinVerHash.set("4.0.950", "Windows 95");
//	WinVerHash.set("4.0.1381", "Windows NT Workstation 4.0");
//	WinVerHash.set("4.1.1998", "Windows 98");
//	WinVerHash.set("4.1.2222", "Windows 98 Second Edition");
//	WinVerHash.set("4.90.3000", "Windows Me");
//	WinVerHash.set("5.0.2195", "Windows 2000 Professional");
//	WinVerHash.set("5.1.2600", "Windows XP");
//	WinVerHash.set("5.2.3790", "Windows Server 2003 R2 SP2");
//	WinVerHash.set("6.0.6000", "Windows Vista");
//	WinVerHash.set("6.1.7600", "Windows 7");
//	WinVerHash.set("6.1.7601", "Windows 7 SP1");
////	The next two lines won't work in Windows 7 UAC.
////	var WinVer = Application.WShell.RegRead("HKLM\\SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\CurrentVersion");
////	var WinBuild = Application.WShell.RegRead("HKLM\\SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\CurrentBuild");
////	var WinVerNum = WinVer + "." + WinBuild;
////	...so how do we get the windows version?

try {
	Application.TempPath = Application.FSO.GetAbsolutePathName(Application.FSO.GetSpecialFolder(2));
}
catch (e) {
	Application.TempPath = "H:";
}

//var sBatCode = "ECHO OFF\r\nFOR /F \"tokens=*\" %%A IN ('VER') DO FOR %%B IN (%%~A) DO FOR /F \"delims=]\" %%C IN (\"%%~B\") DO echo %%C > " + Application.TempPath+ "\\winver.txt\r\nECHO ON";
//var oTSWinVer = Application.FSO.CreateTextFile(Application.TempPath + "\\winver.bat");
//oTSWinVer.WriteLine(sBatCode);
//oTSWinVer.Close();
//Application.Shell(Application.TempPath + "\\winver.bat");
//
//oTSWinVer = Application.FSO.OpenTextFile(Application.TempPath + "\\winver.txt");
//var WinVerNum = oTSWinVer.ReadLine().trim();
//oTSWinVer.Close();

try {
	//Application.DocPath = Application.FSO.GetFolder(Application.TempPath).ParentFolder.ParentFolder.Path + "\\My Documents";
	Application.DocPath = "H:";
	//if (!Application.FSO.FolderExists(Application.DocPath)) {
	//	//	maybe it's Windows 7
	//	Application.DocPath = Application.FSO.GetFolder(Application.TempPath).ParentFolder.ParentFolder.ParentFolder.Path + "\\My Documents";
	//}
	//if (!Application.FSO.FolderExists(Application.DocPath)) {
	//	Application.DocPath = "H:";
	//}
	//Application.Desktop = Application.FSO.GetFolder(Application.TempPath).ParentFolder.ParentFolder.Path + "\\Desktop";
	//if (!Application.FSO.FolderExists(Application.Desktop)) {
	//	//	maybe it's Windows 7
	//	Application.Desktop = Application.FSO.GetFolder(Application.TempPath).ParentFolder.ParentFolder.ParentFolder.Path + "\\Desktop";
	//}
	//if (!Application.FSO.FolderExists(Application.Desktop)) {
	//	Application.Desktop = "H:";
	//}
}
catch (e) {
	Application.DocPath = "C:\\AAWork";
	//Application.Desktop = "H:";
}



Object.extend(Application, {
	IOMode:  {
		ForReading: 1,
		ForWriting: 2,
		ForAppending: 8
	},
	
	//WindowsVersionNumber:  WinVerNum,
	//WindowsVersionName:  WinVerHash.get(WinVerNum),
	
	WinPath: Application.FSO.GetAbsolutePathName(Application.FSO.GetSpecialFolder(0)),
	UserName: Application.Net.UserName,
	ComputerName: Application.Net.ComputerName,

	WeekdayName: ["Sunday", "Monday", "Tuesday", "Wednesday","Thursday", "Friday", "Saturday"],
	WeekdayAbbrev: ["Sun", "Mon", "Tue", "Wed","Thu", "Fri", "Sat"],
	MonthName: ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"],
	MonthAbbrev: ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"],
	ShapeType: ["placeholder","bqLine","bqRectangle","bqRoundRectangle","bqOval","bqTextLabel","bqPicture","bqEmbeddedSection","bqHorizontalLine","bqVerticalLine","bqDropDown","bqListBox","bqTextBox","bqRadioButton","bqCheckBox","bqButton","bqEmbeddedBrowser","bqHyperLink"],
	
	Font: Class.create({
		initialize: function(name, size, style, effect, color) {
			_name = "Arial";
			_size = 10;
			_style = 0;
			_effect = 0;
			_color = 0;
			if (typeof name == "string") _name = name;
			if (typeof size == "number" && size > 7 && size < 73) _size = size;
			if (typeof effect == "number" && effect > -1 && effect < 7) _effect = effect;
			if (typeof style == "number" && style > -1 && style < 5) _style = style;
			
			//	color must be a number between 0 and 16777215
			//	or a string that equates to a hexadecimal value between "000000" and "ffffff"
			if (typeof color == "number") {
				_color = color;
			}
			else if (typeof color == "string") {
				var re = /[^A-Fa-f0-9]/gi;
				var a = color.match(re);
				if (a == null && color.length <= 6) {
					//	the string is valid
					_color = parseInt(color, 16);
				}
			}
		},
		
		Name: function(s) {
			if (typeof s == "string") {
				_name = s;
			}
			return _name;
		},
		Size: function(n) {
			if (typeof n == "number") {
				if (n > 7 && n < 73) {
					_size = n;
				}
			}
			return _size;
		},
		Style: function(n) {
			if (typeof n == "number") {
				if (n > -1 && n < 5) {
					_style = n;
				}
			}
			return _style;
		},
		Effect: function(n) {
			if (typeof n == "number") {
				if (n > -1 && n < 7) {
					_effect = n;
				}
			}
			return _effect;
		},
		Color: function(n) {
			if (typeof n == "number") {
				_color = n;
			}
			else if (typeof n == "string") {
				var re = /[^A-Fa-f0-9]/gi;
				var a = n.match(re);
				if (a == null && n.length <= 6) {
					//	the string is valid
					_color = parseInt(n, 16);
				}
			}
			return _color;
		},
		
		Effects: function() {
			return ["EffectNone", "EffectUnderline", "EffectSubScript", "EffectSuperScript", "EffectStrikeThrough", "EffectOverline", "EffectOverDouble"];
		},
		EffectNone: function() {
			return 0;
		},
		EffectUnderline: function() {
			return 1;
		},
		EffectSubScript: function() {
			return 2;
		},
		EffectSuperScript: function() {
			return 3;
		},
		EffectStrikeThrough: function() {
			return 4;
		},
		EffectOverline: function() {
			return 5;
		},
		EffectOverDouble: function() {
			return 6;
		},
		
		Styles: function() {
			return ["StyleNone", "StyleRegular", "StyleBold", "StyleItalic", "StyleBoldItalic"];
		},
		StyleNone: function() {
			return 0;
		},
		StyleRegular: function() {
			return 1;
		},
		StyleBold: function() {
			return 2;
		},
		StyleItalic: function() {
			return 3;
		},
		StyleBoldItalic: function() {
			return 4;
		},
		
		toString: function() {
			return "Font:  " + _name + ", " + _size.toString() + ", " + _color.toString(16) + ", " + this.Styles()[_style] + ", " + this.Effects()[_effect];
		},
		
		isFont: function() {
			return true;
		},
	}),
	
	CustomCalendar: Class.create({
		initialize: function() {
			_holiday = new Array();		//	array of holidays
		},
		
		isHoliday: function(dt) {
			if (!Date.prototype.isPrototypeOf(dt)) {
				return false;
			}
			if (_holiday.contains(dt.format("mm/dd/yyyy"))) {
				return true;
			}
			else {
				return false;
			}
		},
		
		Add: function (dt) {
			//	add a holiday to the calendar
			if (!Date.prototype.isPrototypeOf(dt)) {
				return false;
			}
			else {
				_holiday.push(dt.format("mm/dd/yyyy"));
				return true;
			}
		},
		
		Holidays: function () {
			//	return an array containing the holidays
			return _holiday;
		}
	}),
	
	Try: {
		these: function() {
			var returnValue;
			
			for (var i = 0, length = arguments.length; i < length; i++) {
				var lambda = arguments[i];
				try {
					returnValue = lambda();
					break;
				} catch (e) { }
			}
			
			return returnValue;
		}
	},
	
	StopWatch: Class.create({
	/*
		Object:		StopWatch
		
		Type:		JavaScript 
		Purpose:	Provides a means of tracking time between events.
		Inputs: 	(no constructor)
		Returns:	StopWatch object
		Notes:		If the start() method is not called, the StopWatch
					start time defaults to when the object was created.
		
		Modification History:
		Date		Developer		Description
		03/18/2008	Keith Stevens	Original
		10/25/2010	Doug Pulse		Converted to JSON
	*/
		initialize:  function () {
			this.start();
		},
	
		start: function () {
			startTS = new Date();
			endTS = new Date();
			this.elapsed = 0;
		},
	
		stop: function () {
			endTS = new Date();
			this.elapsed = (endTS.valueOf() - startTS.valueOf());
		}
	}),
	
	msToTime: function (ms) {
	/*
		Function:	msToTime()
		
		Type:		JavaScript 
		Purpose:	Converts a number of milliseconds into a time-formatted string.
		Inputs: 	ms	number of milliseconds to use
		Returns:	string	a time-formatted string (hh:mm:ss.000)
		
		Modification History:
		Date		Developer		Description
		03/18/2008	Keith Stevens	Original
	*/
		if (ms == NaN) {
			Console.Writeln("msToTime:  " + ms + " is not a number.");
			return ms;
		}
		var sec = Math.floor(ms / 1000);
		ms = ms % 1000;
		t = ms;
		
		var min = Math.floor(sec / 60);
		sec = sec % 60;
		sec = sec.pad(2);
		t = sec + "." + t;
		
		var hr = Math.floor(min / 60);
		min = min % 60;
		min = min.pad(2);
		t = min + ":" + t;
		
		var day = Math.floor(hr / 60);
		hr = hr % 60;
		hr = hr.pad(2);
		t = hr + ":" + t;
		t = day + ":" + t;
		
		return(t);
	},
	
	SortMonthNames: function (a, b) {
		//	function to be used as an argument to Array.sort()
		//	sorts the month names in the correct order for the calendar year
		var retVal = 0;
		var x = a.toProperCase();
		var y = b.toProperCase();
		if (!MonthName.contains(x) && !MonthName.contains(y)) {
			retVal = ((x == y) ? 0 : ((x < y) ? -1 : 1));
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
	},
	
	SortMonthShortNames: function (a, b) {
		//	function to be used as an argument to Array.sort()
		//	sorts the month names in the correct order for the calendar year
		var retVal = 0;
		var x = a.toProperCase().trim();
		var y = b.toProperCase().trim();
		if (!MonthAbbrev.contains(x) && !MonthAbbrev.contains(y)) {
			retVal = ((x == y) ? 0 : ((x < y) ? -1 : 1));
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
	},
	
	ExternalVBScript:  Class.create({
	/*
		Name:		ExternalVBScript
		Purpose:	Provides an object to encapsulate interaction with the operating system through 
					user-defined Windows Scripting Host Visual Basic scripts.
		Inputs: 	esPath		string			The path to the file into which we will write the script.
												default = user's temp path
					esFile		string			The name of the file into which we will write the script.
												default = "TempScript.vbs"
					esScript	string			The script code to write and run.
												This object will tell you if you attempt to run a script 
												before creating it.
					esArgs		string array	An array of name-value pairs of arguments to pass to the script.
					esTimeout	integer			Maximum number of seconds to wait for the script to complete.
					esDebug		boolean			Set to true to write debug information to the the console.
		
		Returns:	An ExternalVBScript object
		
		Notes:		The external script must return a value, even if it just indicates success or failure.
					If you expect to use the results from the Run method, set the timeout long enough to 
					allow your VBScript to complete.  To kick off a long-running script for which no 
					return value is neccessary, set the timeout to 0.
					To keep the Run method from completing before the script has completed its work, have 
					the script create a file named runningMSG in the folder where the script gets written.
					The script must write output to a file named pipe.txt in the folder where the script 
					gets written.
		
		Revision History
		Date		Developer	Description
		12/08/2008	D. Pulse	inagural
		10/28/2010	D. Pulse	converted to JSON
	*/
		initialize:  function(esPath, esFile, esScript, esArgs, esTimeout, esDebug) {
			var slKernel = LoadSharedLibrary("kernel32.dll");
			slKernel.Sleep.Signature = "RI,UI";
			var fsoScript = new JOOLEObject("Scripting.FileSystemObject");
			//var strTempPath = fsoScript.GetAbsolutePathName(fsoScript.GetSpecialFolder(2));	//	user's temp folder
			var strPath = "";
			var strFileName = "";
			var strScript = "";
			var lngTimeout = 60;
			var blnDebug = typeof(g_blnDebug) == "boolean" ? g_blnDebug : false;
			var arrArgs = new Array();
			var strArgs = "";
			
			if (typeof(esPath) == "string") {
				strPath = esPath;
			}
			if (typeof(esFile) == "string") {
				strFileName = esFile;
			}
			if (typeof(esScript) == "string") {
				strScript = esScript;
			}
			if (typeof(esTimeout) == "number") {
				lngTimeout = esTimeout;
			}
			if (typeof(esDebug) == "boolean") {
				blnDebug = esDebug;
			}
			if (esArgs instanceof Array) {
				for (var i = 0; i < esArgs.length; i += 2) {
					if (typeof(esArgs[i]) == "string" && typeof(esArgs[i + 1]) == "string") {
						arrArgs[i] = esArgs[i];
						arrArgs[i + 1] = esArgs[i + 1];
						//" /filter:\"Solutions 2005 Text Files|*.txt|All Files|*.*\" /initdir:\"c:\\aawork\""
						strArgs += " \/" + arrArgs[i] + ":\"" + arrArgs[i + 1] + "\"";
					}
				}
			}
			
			esLogger = new Application.logger();
			esLogger.level(blnDebug ? esLogger.LEVEL.DEBUG : esLogger.LEVEL.ERROR);
			
			//if (strPath.length > 0 && strFileName > 0 && strScript > 0) {
			//	this.Write();
			//	this.Run();
			//	this.getResults();
			//}
			//else {
				if (typeof(esPath) != "string" || strPath.length == 0) 
					strPath = Application.TempPath;
				if (typeof(esFile) != "string" || strFileName.length == 0)
					strFileName = "TempScript.vbs";
				if (typeof(esScript) != "string" || strScript.length == 0)
					strScript = "MsgBox \"The script string was not set.\"";
			//}
		},
		
		Debug: function (blnVal) {
			if (typeof(blnVal) == "boolean")
				blnDebug = blnVal;
		},
		
		setTimeout: function (lngVal) {
			if (typeof(lngVal) == "number")
				lngTimeout = lngVal;
		},
		
		getTimeout: function () {
			return lngTimeout;
		},
		
		setScript: function (strVal) {
			if (typeof(strVal) == "string" && strVal.length > 0) 
				strScript = strVal;
		},
		
		getScript: function () {
			return strScript;
		},
		
		setArgs: function (arrVal) {
			if (arrVal instanceof Array){
				strArgs = "";
				for (var i = 0; i < esArgs.length; i += 2) {
					if (typeof(esArgs[i]) == "string" && typeof(esArgs[i + 1]) == "string") {
						arrArgs[i] = esArgs[i];
						arrArgs[i + 1] = esArgs[i + 1];
						//" /filter:\"Solutions 2005 Text Files|*.txt|All Files|*.*\" /initdir:\"c:\\aawork\""
						strArgs += " \/" + arrArgs[i] + ":\"" + arrArgs[i + 1] + "\"";
					}
				}
			}
		},
		
		getArgs: function () {
			return arrArgs;
		},
		
		getArgString: function () {
			return strArgs;
		},
		
		setPath: function (strVal) {
			if (typeof(strVal) == "string" && strVal.length > 0) 
				strPath = strVal;
		},
		
		getPath: function () {
			return strPath;
		},
		
		setFileName: function (strVal) {
			if (typeof(strVal) == "string" && strVal.length > 0) 
				strFileName = strVal;
		},
		
		getFileName: function () {
			return strFileName;
		},
		
		Reset: function () {
			esLogger.debug("     +++++> ENTERING FUNCTION: ExternalVBScript.Write Script");
			if (fsoScript.FileExists(strPath + "\\" + strFileName)) 
				fsoScript.DeleteFile(strPath + "\\" + strFileName);
			if (fsoScript.FileExists(strPath + "\\pipe.txt")) 
				fsoScript.DeleteFile(strPath + "\\pipe.txt");
			if (fsoScript.FileExists(strPath + "\\runningMSG")) 
				fsoScript.DeleteFile(strPath + "\\runningMSG");
			
			strPath = "";
			strFileName = "";
			strScript = "";
			lngTimeout = 60;
			blnDebug = typeof(g_blnDebug) == "boolean" ? g_blnDebug : false;
			arrArgs = new Array();
			strArgs = "";
			esLogger.debug("     <----- EXITING FUNCTION: ExternalVBScript.Write Script");
		},
		
		WriteScript: function (l_strPath, l_strFile, l_strScript) {
		/*
			Function:	WriteScript()
			Type:		JavaScript
			Purpose:	Writes a Windows script to a file.
			
			Inputs:		l_strPath	string	Temporarily overrides strPath.
						l_strFile	string	Temporarily overrides strFileName.
						l_strScript	string	Temporarily overrides strScript.
			
			Returns:	None
			
			Modification History:
			Date		Developer		Description
						Gary McCool		Original
			12/8/2008	Doug Pulse		Changed from kernel32.Sleep to looping until the file exists.
		*/
			esLogger.debug("     +++++> ENTERING FUNCTION: ExternalVBScript.Write Script");
			
			//	if arguments are missing, use values from the object properties
			if (typeof(l_strPath) != "string" || l_strPath.length == 0)
				l_strPath = strPath;
			if (typeof(l_strFile) != "string" || l_strFile.length == 0)
				l_strFile = strFileName;
			if (typeof(l_strScript) != "string" || l_strScript.length == 0)
				l_strScript = strScript;
			
			
			// Write Script 
			var l_fsoHandle = fsoScript.CreateTextFile(l_strPath + "\\" + l_strFile);
			l_fsoHandle.WriteLine(l_strScript);
			l_fsoHandle.Close();
			
			// the correct way to do this is with checking for the existance or abscence of a file
			// that indicates the the file has been written.  This is necessary due to latency.
			DoEvents();
			var k = (new Date()).valueOf();
						
			while (!fsoScript.FileExists(l_strPath + "\\" + l_strFile)) {
				DoEvents();
				if ((new Date()).valueOf() - k > 5000) {
					esLogger.error("It took too long to write the script.  If the problem persists, please contact technical support.");
					break;
				}
			};
			
			esLogger.debug("     <----- EXITING FUNCTION: ExternalVBScript.Write Script");
		},
		
		RunScript: function (l_strPath, l_strFile, l_arrArgs) {
		/*
			Function:	RunScript()
			Type:		JavaScript 
			Purpose:	Runs a Windows script.
			
			Inputs:		l_strPath	string	Temporarily overrides strPath.
						l_strFile	string	Temporarily overrides strFileName.
						l_arrArgs	string	Temporarily overrides strFileName.
			
			Returns:	None
			
			Modification History:
			Date		Developer		Description
						Gary McCool		Original
			12/8/2008	Doug Pulse		Changed from kernel32.Sleep to looping until the file exists.
		*/
			var strScriptOutput = "";
			
			esLogger.debug("     +++++> ENTERING FUNCTION:  ExternalVBScript.RunScript");
			
			//	if arguments are missing, use values from the object properties
			if (typeof(l_strPath) != "string" || l_strPath.length == 0)
				l_strPath = strPath;
			if (typeof(l_strFile) != "string" || l_strFile.length == 0)
				l_strFile = strFileName;
			var strPipeFile = "pipe.txt";
			
			if (fsoScript.FileExists(l_strPath + "\\" + l_strFile)) {
				esLogger.debug(l_strPath + "\\" + l_strFile + strArgs);
				Application.Shell("wscript", l_strPath + "\\" + l_strFile + strArgs);
				DoEvents();
				slKernel.Sleep(100);
				var k = (new Date()).valueOf();
				esLogger.debug("A");
				while (fsoScript.FileExists(l_strPath + "\\runningMSG")) {
					DoEvents();
					if ( ((new Date()).valueOf() - k) > (lngTimeout * 1000) ) {
						esLogger.error("         It took too long for the user to select the file.");
						break;
					}
					DoEvents();
				};
				esLogger.debug("C");
			}
			else {
				esLogger.error(l_strPath + "\\" + l_strFile + " doesn't exist.");
			};
			
			// get filename from text file.

			var l_fsoInputFile = fsoScript.GetFile(l_strPath + "\\" + strPipeFile);
			esLogger.debug("1");
			var l_fsoInputStream = l_fsoInputFile.OpenAsTextStream(1);
			esLogger.debug("2");
			while (!l_fsoInputStream.AtEndOfStream) {
				strScriptOutput = l_fsoInputStream.ReadLine();
			};
			l_fsoInputStream.Close();
			esLogger.debug("D");
			
			esLogger.debug("     <----- EXITING FUNCTION:  ExternalVBScript.RunScript");
			
			return strScriptOutput;
		}
	})
});


Application.logger = function (level, target, verbose) {
	var _levels = [0, 1, 2, 3, 4, 5, 6, 7], 
		LEVEL = {
				ALL:  0,
				TRACE:  1,
				DEBUG:  2,
				INFO:  3,
				WARN:  4,
				ERROR:  5,
				FATAL:  6,
				OFF:  7
			};
	
	switch (typeof level) {
		case "number" :
			if (_levels.contains(level)) {
				break;
			}
		default :
			level = LEVEL.OFF;
			break;
	}
	switch (typeof target) {
		case "string" :
			a = target.split("\\");
			if (target.toLowerCase() == "console") {
				//Console.Writeln("console");
				target = target.toLowerCase();
			}
			else if (target.toLowerCase() == "alert") {
				target = target.toLowerCase();
			}
			else if (a.length > 1) {
				//Console.Writeln("file");
				if (Application.FSO.FileExists(target)) {
					//ts = Application.FSO.OpenTextFile(0, true, Application.IOMode.ForAppending, target);
					//ts.Close();
				}
				else if (Application.FSO.FolderExists(a.slice(0, -1).join("\\"))) {
					ts = Application.FSO.CreateTextFile(target);
					ts.Close();
				}
			}
			else if (a.length == 1) {
				target = Application.DocPath + "\\" + target;
			}
			else {
				target = "console";
			}
			break;
		case "object" :
			if (typeof target.Type == "undefined") {
				target = "console";
			}
			else {
				if ((new Array(bqTextBox, bqTextLabel, bqDropDown, bqListBox)).contains(target.Type)) {
					//	This is a valid object
				}
				else {
					target = "console";
				}
			}
			break;
		default :
			target = "console";
			break;
	}
	
	var _target = target, 
		_level = level, 
		_verbose = true;
		
	switch (typeof verbose) {
		case "boolean" :
			_verbose = verbose;
			break;
		default :
			_verbose = true;
			break;
	}
	
	var _log = function (level, msg) {
	/*
		purpose:		writes output to the log target
		return values:	0	log entry written
						1	log entry not written
						2	invalid level -- abort
						3	invalid message -- empty log entry written
						4	unhandled error
	*/
		if (_levels.contains(level)) {
			var sMsg = "";
			var r = 4;
			if (typeof msg == "string") {
				sMsg = msg;
			}
			else {
				try {
					sMsg = msg.toString();
				}
				catch (e) {
					sMsg = ""
					r = 3;
				}
			}
			//Console.Writeln(level + "    " + _level);
			if (level >= _level) {
				s = (new Date()).format("mm/dd/yyyy hh:nn:ss") + "\t";
				if (_verbose) {
					s += Application.UserName + "\t";
					s += Application.ComputerName + "\t";
					s += Application.AppType + "\r\n";
					s += ActiveDocument.Path + "\r\n";
					s += ActiveDocument.URL + "\r\n";
				}
				s += sMsg;
				if (_verbose) {
					s += "\r\n";
				}
				
				if (_target == "console") {
					Console.Writeln(s);
				}
				else if (_target == "alert") {
					Alert(s);
				}
				else if (typeof _target == "object") {
					switch (_target.Type) {
						case bqTextBox :
						case bqTextLabel :
							_target.Text += s + "\r\n"
							break;
						case bqDropDown :
						case bqListBox :
							_target.Add(s);
							break;
						default :
							break;
					}
				}
				else {
					try {
						ts = Application.FSO.OpenTextFile(0, true, Application.IOMode.ForAppending, _target);
						ts.WriteLine(s);
						ts.Close();
					}
					catch (e) {
						Console.Writeln(e.toString());
						Console.Writeln("Unable to write log entry to " + _target + ".");
					}
				}
				r = 0;
			}
			else {
				r = 1;
			}
		}
		else {
			r = 2;
		}
		return r;
	}, 
	
	Level = function (level) {
		if (level != undefined) {
			if (_levels.contains(level)) _level = level;
		}
		return _level;
	},
	
	Target = function (target) {
		if (target != undefined) {
			switch (typeof target) {
				case "string" :
					a = target.split("\\");
					if (target.toLowerCase() == "console") {
						target = "console";
						break;
					}
					else if (target.toLowerCase() == "alert") {
						target = "alert";
						break;
					}
					else if (a.length > 1) {
						if (Application.FSO.FileExists(target)) {
							//ts = Application.FSO.OpenTextFile(target, ForAppending);
							//ts.Close();
							break;
						}
						else if (Application.FSO.FolderExists(a.slice(0, -1).join("\\"))) {
							ts = Application.FSO.CreateTextFile(target);
							ts.Close();
							break;
						}
					}
					else if (a.length == 1) {
						target = Application.DocPath + "\\" + target;
						break;
					}
				case "object" :
					if (typeof target.Type == "undefined") {
						target = "console";
					}
					else {
						if ((new Array(bqTextBox, bqTextLabel, bqDropDown, bqListBox)).contains(target.Type)) {
							//	This is a valid object
						}
						else {
							target = "console";
						}
					}
					break;
				default :
					break;
			}
			_target = target;
		}
		
		return _target;
	},
	
	Verbose = function (v) {
		if (v != undefined) {
			if (typeof v == "boolean") _verbose = v;
		}
		return _verbose;
	},
	
	trace = function (msg) {
		return _log(LEVEL.TRACE, msg);
	},
	
	debug = function (msg) {
		return _log(LEVEL.DEBUG, msg);
	},
	
	info = function (msg) {
		return _log(LEVEL.INFO, msg);
	},
	
	warn = function (msg) {
		return _log(LEVEL.WARN, msg);
	},
	
	error = function (msg) {
		return _log(LEVEL.ERROR, msg);
	},
	
	fatal = function (msg) {
		return _log(LEVEL.FATAL, msg);
	};
	
	return {
		LEVEL: LEVEL, 
		Level: Level, 
		Target: Target, 
		Verbose: Verbose, 
		trace: trace, 
		debug: debug, 
		info: info, 
		warn: warn, 
		error: error, 
		fatal: fatal
	}
};

//	This was removed because it may not be appropriate everywhere the library is used.
//	//	if the file is loaded in the web client, don't prompt to save when the document is closed
//	if ((Application.Type == bqAppTypePlugInClient) || (Application.Type == bqAppTypeThinClient))
//		ActiveDocument.PromptToSave = false;



var sUUID = "";
var sReposPath = "";
if (Application.Hyperion.Type.Insight || Application.Hyperion.Type.Scheduler) {
	if (ActiveDocument.URL.indexOf("&DOC_UUID") == -1) {
		//	document is open in insight, but not from the Workspace
		sUUID = "{UUID not found}";
		sReposPath = decodeURIComponent(ActiveDocument.URL.replace(/file:\/\//gi, ""));
		sFolder = sReposPath.substring(0, sReposPath.lastIndexOf("\\") + 1);
	}
	else {
		sUUID = ActiveDocument.URL.substring(ActiveDocument.URL.indexOf("&DOC_UUID") + 10, ActiveDocument.URL.indexOf("&DOC_VERSION"));
		sReposPath = decodeURIComponent(ActiveDocument.URL.substring(ActiveDocument.URL.indexOf("&repository_path") + 17, ActiveDocument.URL.indexOf("&repository_token")));
		sFolder = sReposPath.substring(0, sReposPath.lastIndexOf("/") + 1);
	}
}
else {
	sUUID = "{UUID not found}";
	sReposPath = ActiveDocument.Path;
	sFolder = sReposPath.substring(0, sReposPath.lastIndexOf("\\") + 1);
}

Object.extend(Application.ActiveDocument, {
	UUID:  sUUID, 
	ReposPath:  sReposPath, 
	Folder: sFolder, 
	
	ClearResults: function (arrQuery, arrProcess) {
		//	Each element arrQuery and arrProcess refers to the same query.
		//	arrQuery contains the query objects.
		//	arrProcess contains boolean values telling whether or not to process the query with the same element number.
		//	Example:
		//		arrProcess[0] = true;	//	process arrQuery[0]
		//		arrProcess[1] = false;	//	don't process arrQuery[1]
		for (var i = 0; i < arrQuery.length; i++) {
			if (typeof arrProcess == "undefined") {
				arrQuery[i].Limits["Id"].Ignore = false;
				arrQuery[i].Process();
				arrQuery[i].Limits["Id"].Ignore = true;
			}
			else if (arrProcess[i]) {
				arrQuery[i].Limits["Id"].Ignore = false;
				arrQuery[i].Process();
				arrQuery[i].Limits["Id"].Ignore = true;
			}
		}
	},	//	ClearResults
	
	ExportData: function (oDash, sResultPrefix, sMessageTitle) {
		//	Purpose:  export the data from the report to pdf or from the results to excel
		var d = new Date();
		var dfile = new String("");
		var iOutFormat, sSection;
		var sBaseName = oDash.Name;
		var sOut;
		
		if (typeof sMessageTitle != "string") sMessageTitle = "";
		
		if (oDash.ActiveTab) {
			//	to accommodate multiple levels of tabs
			sBaseName += " - " + oDash.ActiveTab;
		}
		
		if (oDash.FPType) {
			//	to allow file naming based on the type of fiscal period
			sBaseName = sBaseName.replace(/ Year/gi, " " + oDash.FPType + " Year");
		}
		
		dfile = sBaseName + "_" + d.format("yyyymmddhhnnss");
		
		var iOut = Alert("What do you want to export?", sMessageTitle, "Printable Document", "Data to Spreadsheet", "Cancel");
		switch (iOut) {
			case 1 :
				dfile += ".pdf";
				iOutFormat = bqExportFormatPDF;
				
				if (oDash.ActiveTab) {
					sSection = sBaseName;
				}
				else {
					sSection = sBaseName.replace(/ /gi, "  ");
				}
				
				ActiveDocument.Sections[sSection].Export(ActiveDocument.g_strUserTempPath + "\\" + dfile, iOutFormat, true, false);
				Shell(g_strWinPath + "\\explorer.exe", "\"" + ActiveDocument.g_strUserTempPath + "\\" + dfile + "\"");
				sOut = ActiveDocument.g_strUserTempPath + "\\" + dfile
				break;
				
			case 2 :
				sSection = sResultPrefix + "-" + sBaseName;
				sOut = ActiveDocument.Sections[sSection].ExportToExcel(ActiveDocument.g_strUserTempPath, dfile, true);
				
				break;
				
			case 3 : 
				//	user cancelled
				break;
				
			default :
				break;
		}
		return sOut;
	}, 	//	ExportData
	
	SectionExists:  function (strSectionName) {
	/*
		Method:	SectionExists()
		Type:		JavaScript 
		Purpose:	Determine whether the named section exists in the ActiveDocument.
		Inputs:		strSectionName		string	name of the object for which to search
		Returns:	boolean
		
		Modification History:                                             
		Date		Developer	Description
		10/29/2008	Doug Pulse	Original
	*/
		var strTemp = "";
		var blnRetVal = false;
		try {
			//	see if strSectionName is the name of a section object
			strTemp = ActiveDocument.Sections.Item(strSectionName).Name;
		}
		catch (e) {
			//	strSectionName is not the name of a section object
			Control.Writeln(strSectionName + " is not the name of a Section object.");
			strTemp = "";
		}
		if (strTemp == "") {
			blnRetVal = false;
		}
		else {
			blnRetVal = true;
		}
		return (blnRetVal);
	}, 	//SectionExists
	
	OutputScripts:  function(strOutputFile) {
		//	export all scripts from the active document to a text file
		var strOutput = "";
		var strScript = "";
		if (strOutputFile == undefined) {
			strOutputFile = Application.TempPath + "\\scriptoutput.txt";
		}
		if (typeof strOutputFile != "string") {
			strOutputFile = Application.TempPath + "\\scriptoutput.txt";
		}
		
		//	Cycle through all of the event scripts in the active document.
		for (var i = 1; i <= ActiveDocument.EventScripts.Count; i++) {
			strScriptName = ActiveDocument.EventScripts.Item(i).Name;
			strScript = "/* new script */\r\n";
			strScript += "//script:     " + strScriptName + "\r\n";
			strScript += "//object:     ActiveDocument.EventScripts[\"" + strScriptName + "\"].Script = \r\n";
			strScript += ActiveDocument.EventScripts.Item(i).Script;
			strScript += "\r\n";
			strOutput += strScript;
		}
		
		//	Cycle through all of the sections in the ActiveDocument.
		for (var i = 1; i <= ActiveDocument.Sections.Count; i++) {
			//	Is it a dashboard?
			if (ActiveDocument.Sections.Item(i).Type == bqDashboard) {
				//	Cycle through all of the event scripts in the dashboard.
				strDashboardName = ActiveDocument.Sections.Item(i).Name;
				for (var j = 1; j <= ActiveDocument.Sections.Item(i).EventScripts.Count; j++) {
					strScriptName = ActiveDocument.Sections.Item(i).EventScripts.Item(j).Name;
					strScript = "/* new script */\r\n";
					strScript += "//dashboard:  " + strDashboardName + "\r\n";
					strScript += "//script:     " + strScriptName + "\r\n";
					strScript += "//object:     ActiveDocument.Sections[\"" + strDashboardName + "\"].EventScripts[\"" + strScriptName + "\"].Script = \r\n";
					strScript += ActiveDocument.Sections.Item(i).EventScripts.Item(j).Script;
					strScript += "\r\n";
					strOutput += strScript;
				}
				//	Cycle through all of the shapes in the dashboard.
				for (j = 1; j <= ActiveDocument.Sections.Item(i).Shapes.Count; j++) {
					//	Cycle through all of the event scripts in the shape.
					strShapeName = ActiveDocument.Sections.Item(i).Shapes.Item(j).Name;
					for (var k = 1; k <= ActiveDocument.Sections.Item(i).Shapes.Item(j).EventScripts.Count; k++) {
						strScriptName = ActiveDocument.Sections.Item(i).Shapes.Item(j).EventScripts.Item(k).Name;
						strScript = "/* new script */\r\n";
						strScript += "//dashboard:  " + strDashboardName + "\r\n";
						strScript += "//shape:      " + strShapeName + "\r\n";
						strScript += "//script:     " + strScriptName + "\r\n";
						strScript += "//object:     ActiveDocument.Sections[\"" + strDashboardName + "\"].Shapes[\"" + strShapeName + "\"].EventScripts[\"" + strScriptName + "\"].Script = \r\n";
						strScript += ActiveDocument.Sections.Item(i).Shapes.Item(j).EventScripts.Item(k).Script;
						strScript += "\r\n";
						strOutput += strScript;
					}
				}
			}
		}
		
		try {
			var oTS = Application.FSO.CreateTextFile(false, true, strOutputFile);
			oTS.WriteLine(strOutput);
			oTS.Close();
		}
		catch (e) {
			Console.Writeln("Scripts could not be output to " + strOutputFile);
		}
		return strOutput;
	}, 	//	OutputScripts
	
	InputScripts:  function (strInputFile) {
		//	load scripts from a text file into the ActiveDocument
		var strInput = "";
		var strScript = "";
		var strDashboardName = "";
		var strShapeName = "";
		var strScriptName = "";
		var objDashboardName;
		var objShapeName;
		var objScriptName;
		var strObjectName = "";
		
		//	validate input
		if (strInputFile == undefined) {
			return "Input file name is required";
		}
		if (typeof strInputFile != "string") {
			return "Input file name must be a string";
		}
		if (!Application.FSO.FileExists(strInputFile)) {
			return strInputFile + " doesn't exist."
		}
		
		var oTS = Application.FSO.OpenTextFile(0, false, 1, strInputFile);
		while (!oTS.AtEndOfStream) {
			strInput = oTS.ReadLine();
			if (strInput == "\/* new script *\/") {
				if (strScriptName != "") {
					//	load the previous script
					try {
						objScript.Script = strScript.substr(0, strScript.length - 2);
					}
					catch (e) {
						Console.Writeln(strObjectName + ".Script could not be set.");
					}
				}
				//	read the new script
				strDashboardName = "";
				objDashboard = new Object();
				strShapeName = "";
				objShape = new Object();
				strScriptName = "";
				objScript = new Object();
				strObjectName = "ActiveDocument";
				
				while (strInput.indexOf("\/\/object") == -1) {
					strInput = oTS.ReadLine();
					if (strInput.substr(0, 14) == "//dashboard:  ") {
						strDashboardName = strInput.substr(14);
						strObjectName += ".Sections[\"" + strDashboardName + "\"]";
						try {
							objDashboard = ActiveDocument.Sections[strDashboardName];
						}
						catch (e) {
							Console.Writeln(strObjectName + " does not exist.");
						}
					}
					if (strInput.substr(0, 14) == "//shape:      ") {
						strShapeName = strInput.substr(14);
						strObjectName += ".Shapes[\"" + strShapeName + "\"]";
						try {
							objShape = objDashboard.Shapes[strShapeName];
						}
						catch (e) {
							Console.Writeln(strObjectName + " does not exist.");
						}
					}
					if (strInput.substr(0, 14) == "//script:     ") {
						strScriptName = strInput.substr(14);
						strObjectName += ".EventScripts[\"" + strDashboardName + "\"]";
						if (strDashboardName == "" && strShapeName == "") {
							//	it's an ActiveDocument EventScript
							try {
								objScript = ActiveDocument.EventScripts[strScriptName];
							}
							catch (e) {
								Console.Writeln(strObjectName + " does not exist.");
							}
						}
						else {
							if (strDashboardName != "" && strShapeName == "") {
								//	it's a dashboard eventscript
								try {
									objScript = objDashboard.EventScripts[strScriptName];
								}
								catch (e) {
									Console.Writeln(strObjectName + " does not exist.");
								}
							}
							else {
								//	it's a shape eventscript
								try {
									objScript = objShape.EventScripts[strScriptName];
								}
								catch (e) {
									Console.Writeln(strObjectName + " does not exist.");
								}
							}
						}
					}
				}
				strScript = "";
			}
			else {
				strScript += strInput + "\r\n";
			}
		}
		//	load the last script
		try {
			objScript.Script = strScript.substr(0, strScript.length - 2);
		}
		catch (e) {
			Console.Writeln(strObjectName + ".Script could not be set.");
		}
		
		oTS.Close( );
	}, 	//	InputScripts
	
	SynchModels: function () {
		for (var i = 1, q, j, blnDone; i <= ActiveDocument.Sections.Count; i++) {
			q = ActiveDocument.Sections.Item(i)
			if (q.Type == bqQuery) {
				Console.Write(ActiveDocument.Sections.Item(i).Name);
				j = 0;
				blnDone = false;
				while (j < 5 && !blnDone) {
					Console.Write(".");
					j++;
					try {
						q.DataModel.SyncWithDatabase();
					}
					catch (e) {
						blnDone = true;
					}
					DoEvents();
				}
				Console.Writeln("");
			}
		}
	}, 	//	SynchModels
	
	createEmail: function (arrTo, strSubject, strMessage, arrAttach, blnIncludeLink) {
	/*
		Function:	createEmail()
		
		Type:		JavaScript 
		Purpose:	Create an e-mail with a link to the current bqy file.
					Displays the message ready to be sent.
		Inputs:		arrTo			array of strings	(optional)	e-mail addresses to include in the To: box
					strSubject		string				(optional)	subject of the message
					strMessage		string				(optional)	additional message to include
					arrAttach		array of strings	(optional)	paths to files to attach
					blnIncludeLink	boolean				(optional)	Should the message include a hyperlink to the document on the Workspace?  Defaults to true.
		
		Returns:	None
		
		Dependencies:	Microsoft Outlook
		
		Modification History:                                             
		Date		Developer	Description
		7/25/2008	Doug Pulse	Original
		1/4/2012	Doug Pulse	Added blnIncludeLink
	*/
		//	check input to ensure initial values are good
		//	bad input will be ignored
		var strTo = "";
		if (typeof arrTo == "undefined") {
			strTo = "";
		}
		else if (arrTo instanceof Array) {
			for (var a in arrTo) {
				if (typeof arrTo[a] == "string") strTo += ";" + arrTo[a];
			}
			strTo = strTo.substr(1);
		}
		else {
			strTo = "";
		}
		if (typeof strSubject != "string") strSubject = "Link to Report";
		if (typeof strMessage != "string") strMessage = "Please follow this link to the report:\n\n";
		if (typeof blnIncludeLink != "boolean") blnIncludeLink = true;
		
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
		strSmartCut = strPreCut + strSmartCut + strPostCut;
		
		if (blnIncludeLink) strMessage += strSmartCut + "\n\n";
		
		var olApp = new JOOLEObject("Outlook.Application");
		var olNote = olApp.CreateItem(0);
		olNote.To = strTo;
		olNote.Subject = strSubject;
		olNote.Body = strMessage;
		
		if (arrAttach instanceof Array) {
			for (var a in arrAttach) {
				if (typeof arrAttach[a] == "string") {
					//	if the path doesn't exist, just don't attach the file
					try {
						olNote.Attachments.Add (arrAttach[a]);
					}
					catch (e) { }
				}
			}
		}
		//olNote.Send;
		olNote.Display();
	},	//	createEmail
	
	loadObjectMethods:  function () {
		for (var i = 1; i <= ActiveDocument.Sections.Count; i++) {
			var oSec = ActiveDocument.Sections.Item(i);
			
			//	In the all queries in the ActiveDocument, create a method to automatically import a sql file and set the column names
			if (oSec.Type == bqQuery) {
				oSec.ReImportSQL = function (sFileName) {
					var iOut = 1;
					
					//	not applicable if the data model has topics
					if (!this.DataModel.Topics.Count) {
						var FOR_READING = 1;
						var oFSO = new JOOLEObject("Scripting.FileSystemObject");
						
						if (oFSO.FileExists(sFileName)) {
							var oTS = oFSO.OpenTextFile(FOR_READING, sFileName);
							if (!oTS.AtEndOfStream) {
								//	read the sql file
								var sContent = oTS.ReadAll();
								//	store the whole sql statement
								var sSQL = sContent;
								//	extract the column aliases
								//	The query must be in a specific format
								//		select *
								//		from (
								//		-- query goes here
								//		) q
								//		--cols: n
								//		(
								//		--column aliases go here
								//		)
								//
								//	The column aliases must be in a specific format.
								//		line 1:		--cols: n		where n = number of columns
								//		line 2: 	(
								//		line 3:		[column 1 alias]
								//		line 4:		, [column 2 alias]
								//		line n+2:	, [column n alias]
								//		line n+3:	)
								var sCols = sContent.substr(sContent.indexOf("\r\n--cols: ") + 2);
								//	put the column aliases into an array
								//	The array should have n+3 elements
								var arrCols = sCols.split("\r\n");
								//	get the number of columns from the comment
								var iColsGiven = arrCols[0].split(":")[1] * 1;
								var iColsFound = arrCols.indexOf(")") - arrCols.indexOf("(") - 1;
								if (iColsGiven == iColsFound) {
									var arrCol = arrCols.slice(arrCols.indexOf("(") + 1, arrCols.indexOf(")"));
									for (var i = 0; i < arrCol.length; i++) {
										if (((i == 0 && arrCol[i].substr(0, 1) == "[") || (arrCol[i].substr(0, 3) == ", [")) && (arrCol[i].substr(arrCol[i].length - 1) == "]")) {
											if (i == 0) arrCol[i] = arrCol[i].substr(1);
											else arrCol[i] = arrCol[i].substr(3);
											arrCol[i] = arrCol[i].substr(0, arrCol[i].length - 1);
										}
										else {
											Console.Writeln(this.Name + ".ReImportSQL:  " + sFileName + ":  Invalid column alias for column " + i.toString());
											iOut = 0;
											break;
										}
									}
									if (iOut) {
										//	we have column names
										//	remove the request items
										this.Requests.RemoveAll();
										//	import the SQL
										this.ImportSQLFile(sFileName, iColsFound);
										//	rename the columns
										for (var i = 0; i < arrCol.length; i++) {
											this.Requests[i + 1].DisplayName = arrCol[i];
										}
									}
								}
								else {
									Console.Writeln(this.Name + ".ReImportSQL:  " + sFileName + ":  Columns given = " + iColsGiven.toString() + "\r\nColumns found = " + iColsFound.toString());
									iOut = 0;
								}
							}
							else {
								Console.Writeln(this.Name + ".ReImportSQL:  " + sFileName + ":  File is empty");
								iOut = 0;
							}
						}
						else {
							Console.Writeln(this.Name + ".ReImportSQL:  " + sFileName + ":  File not found");
							iOut = 0;
						}
					}
					else {
						Console.Writeln(this.Name + ".ReImportSQL:  Data model is not empty");
						iOut = 0;
					}
					
					return iOut;
				}
			}
			//	In the all dashboards in the ActiveDocument, create a GetValue, SetValue, GetList, and SetList for each shape (if appropriate).
			if (oSec.Type == bqDashboard) {
				oSec.SaveUserValues = function (sFileName) {
					try {
						if (typeof sFileName == "undefined") {
							sFileName = Application.DocPath + "\\" + ActiveDocument.Name.replace(/ /gi, "_").replace(/[.]/gi, "_").replace(/.bqy$/i, "") + ".qc";
						}
						var oTS = Application.FSO.CreateTextFile(sFileName);
						var sOut = "";
						var hValues = new Application.Hash();
						
						for (var i = 1; i <= this.Shapes.Count; i++) {
							var oShape = this.Shapes.Item(i);
							if ((new Array(bqCheckBox,bqRadioButton,bqTextBox,bqDropDown,bqListBox)).contains(oShape.Type)) {
								hValues.set(oShape.Name, oShape.GetValue());
							}
						}
						
						oTS.WriteLine("{" + ActiveDocument.Name.replace(/ /gi, "_").replace(/[.]/gi, "_").replace(/.bqy$/i, "") + " = {");
						
						for (var i = 0; i < hValues.keys().length; i++) {
							oTS.Write("\t\"" + hValues.keys()[i] + "\": [");
							for (var j = 0; j < hValues.values()[i].length; j++) {
								if (j > 0) oTS.Write(", ");
								oTS.Write("\"" + hValues.values()[i][j].toString() + "\"");
							}
							oTS.Write("]");
							if ((i + 1) < hValues.keys().length) {
								oTS.WriteLine(",");
							}
							else {
								oTS.WriteLine("");
							}
						}
						oTS.WriteLine("}}");
						oTS.Close();
					}
					catch (e) {
						Console.Writeln(e.toString());
						oTS.Close();
					}
				};	//	SaveUserValues
				
				oSec.LoadUserValues = function (sFileName) {
					try {
						if (typeof sFileName == "undefined") sFileName = Application.DocPath + "\\" + ActiveDocument.Name.replace(/ /gi, "_").replace(/[.]/gi, "_").replace(/_bqy$/i, "") + ".qc";
						try {
							var oTS = Application.FSO.OpenTextFile(sFileName);
						}
						catch (e) {
							Console.Writeln("File not found.");
							return 0;
						}
						var sIn = oTS.ReadAll();
						oTS.Close();
						eval(sIn);
						var o = eval(ActiveDocument.Name.replace(/ /gi, "_").replace(/[.]/gi, "_").replace(/_bqy$/i, ""));
						for (var i in o) {
							//	What if the shape doesn't exist?
							try {
								this.Shapes[i].SetValue(o[i]);
								//ActiveDocument.Sections["Selection Criteria"].Shapes[i].SetValue(Expenditures_by_PIN[i]);
							}
							catch (e){
								//	ignore it
							}
						}
					}
					catch (e) {
						Console.Writeln(e.toString());
						oTS.Close();
					}
					
					return 1;
				};	//	LoadUserValues
				
				oSec.ControlExists = function (sControlName) {
					/*
						Method:		Dashboard.ControlExists()
						Type:		JavaScript 
						Purpose:	Determine whether the named section exists in the ActiveDocument.
						Inputs:		sControlName		string	name of the control for which to search
						
						Returns:	boolean
						
						Modification History:                                             
						Date		Developer	Description
						10/29/2008	Doug Pulse	Original
					*/
						var strTemp = "";
						var intType = -1;
						var blnRetVal = false;
						
						//	verify that sControlName is the name of a control in the Dashboard
						try {
							strTemp = this.Shapes[sControlName].Name;
						}
						catch (e) {
							//Console.Writeln(sControlName + " is not a control in " + this.Name + ".");
						}
						
						if (strTemp == "") {
							blnRetVal = false;
						}
						else {
							blnRetVal = true;
						}
						return (blnRetVal);
				};	//	ControlExists
				
				for (var j = 1; j <= oSec.Shapes.Count; j++) {
					var oShape = oSec.Shapes.Item(j);
					
					switch (oShape.Type) {
						case bqRectangle :
						case bqRoundRectangle :
						case bqPicture :
						case bqEmbeddedSection :
						case bqEmbeddedBrowser :
							oShape.maximize = function (s) {
								//	s	type of resize	h = height, w = width, omit for both
								if (typeof s == "string") {
									if (s == "h") this.Placement.Height = ActiveDocument.ContentHeight - this.Placement.YOffset - 20;
									if (s == "w") this.Placement.Width = ActiveDocument.ContentWidth - this.Placement.XOffset - 20;
								}
								else {
									this.Placement.Height = ActiveDocument.ContentHeight - this.Placement.YOffset - 20;
									this.Placement.Width = ActiveDocument.ContentWidth - this.Placement.XOffset - 20;
								}
							};
							break;
						case bqCheckBox :
							oShape.GetValue = function () {
								return(new Array([this.Checked]));
							};
							oShape.SetValue = function (arr) {
								this.Checked = arr[0];
							};
							break;
						case bqRadioButton :
							oShape.GetValue = function () {
								return(new Array([this.Checked]));
							};
							oShape.SetValue = function (arr) {
								this.Checked = arr[0];
							};
							break;
						case bqTextLabel :
							oShape.GetValue = function () {
								return(new Array([this.Text]));
							};
							oShape.SetValue = function (arr) {
								this.Text = arr[0];
							};
							break;
						case bqTextBox :
							oShape.GetValue = function () {
								return(new Array([this.Text]));
							};
							oShape.SetValue = function (arr) {
								this.Text = arr[0];
							};
							break;
						case bqDropDown :
							oShape.GetValue = function () {
								var arr = new Array();
								try { arr = new Array([this.Item(this.SelectedIndex)]) }
								catch (e) { }
								return arr;
							};
							oShape.SetValue = function (arr) {
								for (var k = 1; k <= this.Count; k++) {
									if (this.Item(k) == arr[0]) {
										this.Select(k);
										break;
									}
								}
							};
							oShape.GetList = function () {
								var arr = new Array();
								for (var k = 1; k <= this.Count; k++) {
									arr.push(this.Item(k));
								}
								return(arr);
							};
							oShape.SetList = function (arr) {
								var tTime = new Date();
								var blnGood = false;
								if (Array.prototype.isPrototypeOf(arr)) {
									//	arr is already an array
									blnGood = true;
								}
								else {
									//	arr should be a column in a table or result
									try {
										if (Object.isInteger(arr.ColumnType)) {
											//	arr is a column object
											//	extract the data to an array
											//	eliminate duplicates and sort the array
											arr = arr.toArray().unique(true);
										}
										blnGood = true;
									}
									catch (e) {
										//	arr is not a column
										Console.Writeln("parameter must be a column object or an array");
									}
								}
								
								if (blnGood) {
									this.RemoveAll();
									for (var k = 0; k < arr.length; k++) {
										this.Add(arr[k]);
									
									}
								}
								return new Number(((new Date()).valueOf() - tTime.valueOf()) / 1000);
							};
							break;
						case bqListBox :
							oShape.GetValue = function () {
								var arr = new Array();
								for (var k = 1; k <= this.SelectedList.Count; k++) {
									arr.push(this.SelectedList.Item(k));
								}
								return(arr);
							};
							oShape.SetValue = function (arr) {
								for (var m = 1; m <= this.Count; m++) {
									this.Unselect(m);
									for (var k = 0; k < arr.length; k++) {
										if (this.Item(m) == arr[k]) {
											this.Select(m);
										}
									}
								}
							};
							oShape.GetList = function () {
								var arr = new Array();
								for (var k = 1; k <= this.Count; k++) {
									arr.push(this.Item(k));
								}
								return(arr);
							};
							oShape.SetList = function (arr) {
								var tTime = new Date();
								var blnGood = false;
								if (Array.prototype.isPrototypeOf(arr)) {
									//	arr is already an array
									blnGood = true;
								}
								else {
									//	arr should be a column in a table or result
									try {
										if (Object.isInteger(arr.ColumnType)) {
											//	arr is a column object
											//	extract the data to an array
											//	eliminate duplicates and sort the array
											arr = arr.toArray().unique(true);
										}
										blnGood = true;
									}
									catch (e) {
										//	arr is not a column
										Console.Writeln("parameter must be a column object or an array");
									}
								}
								
								if (blnGood) {
									this.RemoveAll();
									for (var k = 0; k < arr.length; k++) {
										this.Add(arr[k]);
									}
								}
								return new Number(((new Date()).valueOf() - tTime.valueOf()) / 1000);
							};
							break;
						default :
							break;
					}
				}	//	end adding shape methods
			}
			
			if ([bqResults, bqTable, bqPivot].contains(oSec.Type)) {
				//oSec.ExportToExcel = function (strDirectory, strExpFileName, blnOpenWhenDone, blnUnmergeCells) {
				oSec.ExportToExcel = function (strDirectory, strExpFileName, blnOpenWhenDone, blnUnmergeCells, blnUnPivot) {
					/***********************************************************************************
					*	Code copied from http://it.toolbox.com/wiki/index.php/Export_Pivots_and_Results_from_BQY_to_Excel_with_Style
					*	
					*	This function takes the following 4 arguments
					*	Arg Name			Arg Type	Arg Description
					*	strDirectory		String		The directory to export to
					*	strExpFileName		String		The name to use for the exported file
					*									(*exclude file extension)
					*	blnOpenWhenDone		Boolean		indicates whether to open the file after exporting
					*	blnUnmergeCells		Boolean		indicates whether to unmerge merged cells (when exporting from a pivot)
					*	blnUnPivot			Boolean		indicates whether to unpivot the data from a Pivot, putting the column labels before the row labels
					*	
					*	Return value:		the path and name of the file created
					*	
					*	Sample function call:
					*	ExportToExcel("C:\\temp", "Report0207-A", true);
					*	This would create a file named Report0207-A.xls in the temp directory on the 
					*	user's C drive containing the contents of the section object and open the file.
					***********************************************************************************/
					
					//	development note:
					//		use blnUnPivot instead of blnKeepFormatting
					//		Ask the user "how do you want it displayed", "as-is", "unmerged", "unpivotted"
					//		
					
					
					//Console.Writeln("Exporting section to " + strDirectory + "\\" + strExpFileName + ".xls");
					var sOut;
					
					try {
						//**********	validate input
						if (typeof strDirectory == "undefined" || strDirectory == "") throw("NO_DIRECTORY");
						if (typeof strDirectory != "string") throw("BAD_DIRECTORY");
						if (typeof strExpFileName == "undefined" || strExpFileName == "") throw("NO_FILENAME");
						if (typeof strExpFileName != "string") throw("BAD_FILENAME");
						if (typeof blnOpenWhenDone == "undefined") blnOpenWhenDone = false;
						blnOpenWhenDone = new Boolean(blnOpenWhenDone);
						
						if (this.Type == bqPivot) {
							//	If it's not a Pivot, we don't care about the last 2 parameters.
							if (typeof blnUnmergeCells == "undefined" && typeof blnUnPivot == "undefined") {
								var iHow = Alert("How do you want the data exported?", "Export To Excel", "Merged", "Unmerged", "Unpivotted");
								switch (iHow) {
									case 1:
										blnUnmergeCells = false;
										blnUnPivot = false;
										break;
									case 2:
										blnUnmergeCells = true;
										blnUnPivot = false;
										break;
									case 3:
										blnUnmergeCells = false;
										blnUnPivot = true;
										break;
									default:
										throw("UNKNOWN_OPTION");
										break;
								}
							}
							else if (typeof blnUnmergeCells == "undefined" && !blnUnPivot) {
								//	blnUnPivot is defined
								blnUnPivot = new Boolean(blnUnPivot);
								//	If the data is unpivotted, it's already unmerged.
								//	if blnUnPivot is false, ask about unmerging
								if (!blnUnPivot) {
									var iUnMerge = 2;
									
									if (this.Type == bqPivot) {
										iUnMerge = Alert("Do you want the cells unmerged?", "Export to Excel", "Yes", "No");
									}
									
									if (iUnMerge == 1) {
										blnUnmergeCells = true;
									}
									else {
										blnUnmergeCells = false;
									}
								}
							}
							else if (typeof blnUnPivot == "undefined") {
								//	blnUnmergeCells is defined
								blnUnmergeCells = new Boolean(blnUnmergeCells);
								//	If the data is unmerged, it can still be pivotted.
								//	if blnUnmergeCells is true, ask about unpivotting
								if (blnUnmergeCells) {
									var iUnPivot = 2;
									
									if (this.Type == bqPivot) {
										iUnPivot = Alert("Do you want the cells unpivotted?", "Export to Excel", "Yes", "No");
									}
									
									if (iUnPivot == 1) {
										blnUnPivot = true;
									}
									else {
										blnUnPivot = false;
									}
								}
							}
						}
						else {
							blnUnmergeCells = false;
							blnUnPivot = false;
						}
						
						if (blnUnmergeCells) {
							strFileExt = ".xls";
						}
						else {
							strFileExt = ".mhtml";
						}
						
						//**********	Verify The Directory Exists
						if (!Application.FSO.FolderExists(strDirectory)) {
							throw("INVALID_DIRECTORY");
						}
						//**********	ensure the directory name ends in a slash
						if (strDirectory.right(1) != "\\") strDirectory += "\\";
						//Console.Writeln("	folder exists");
						
						var strFile = strDirectory + strExpFileName + strFileExt;
						//Console.Writeln("	file names defined");
						
						//**********	Turn Off Page Breaks
						this.HTMLHorizontalPageBreakEnabled = false;
						this.HTMLVerticalPageBreakEnabled = false;
						
						//**********	If Files Already Exist - Delete Them
						if (Application.FSO.FileExists(strFile)) {
							try {
								Application.FSO.DeleteFile (strFile);
							}
							catch(e) {
								if (e.toString() != "") {
									throw("FILE_STILL_OPEN");
								}
							}
						}
						
						//Console.Writeln("	ready to export");
						
						if (blnUnPivot) {
							//Console.Writeln("keep formatting");
							//	create a table based on the pivot and export it
							var arrET = this.ExportTable();
							arrET[0].ExportToStream(strFile, bqExportFormatOfficeMHTML);
							for (var i = 0; i < arrET.length; i++) {
								arrET[i].Remove();
							}
						}
						else if (blnUnmergeCells) {
							//Console.Writeln("unmerge");
							//	export the Pivot as data for further analysis
							//	remove totals before exporting
							var arrTop = new Array();
							var arrSide = new Array();
							var oTotal;
							
							for (var i = 1; i <= this.TopLabels.Count; i++) {
								for (var j = 1; j = this.TopLabels.Item(i).Totals.Count; j++) {
									arrTop.push([this.TopLabels.Item(i).Name, this.TopLabels.Item(i).Totals.Item(j).DataFunction]);
									this.TopLabels.Item(i).Totals.Item(j).Remove();
								}
							}
							
							for (var i = 1; i <= this.SideLabels.Count; i++) {
								for (var j = 1; j = this.SideLabels.Item(i).Totals.Count; j++) {
									arrSide.push([this.SideLabels.Item(i).Name, this.SideLabels.Item(i).Totals.Item(j).DataFunction]);
									this.SideLabels.Item(i).Totals.Item(j).Remove();
								}
							}
							
							//**********	Export to Excel 97/2003 Format
							this.ExportToStream(strFile, bqExportFormatExcel8);
							//Console.Writeln("	xls file exported");
							
							//	add the totals that were removed, in reverse order to match what was there
							for (j = arrSide.length - 1; j >= 0; j--) {
								for (var i = 1; i <= this.SideLabels.Count; i++) {
									if (arrSide[j][0] == this.SideLabels.Item(i).Name) {
										oTotal = this.SideLabels.Item(i).Totals.Add();
										oTotal.DataFunction = arrSide[j][1];
									}
								}
							}
							
							for (j = arrTop.length - 1; j >= 0; j--) {
								for (var i = 1; i <= this.TopLabels.Count; i++) {
									if (arrTop[j][0] == this.TopLabels.Item(i).Name) {
										oTotal = this.TopLabels.Item(i).Totals.Add();
										oTotal.DataFunction = arrTop[j][1];
									}
								}
							}
						}
						else {
							//**********	Export to MHTML Format
							this.ExportToStream(strFile, bqExportFormatOfficeMHTML);
							//Console.Writeln("	mhtml file exported");
						}
						
						////**********	Get path to EXCEL executable
						//var objExcel = new JOOLEObject("Excel.Application");
						//Console.Writeln("	created Excel object");
						//var strExcel = objExcel.Path + "\\Excel.exe";
						//objExcel.DisplayAlerts = false;
						//objExcel.Quit();
						var strExcel = (new JOOLEObject("Excel.Application")).Path + "\\Excel.exe";
						
						//Console.Writeln(blnOpenWhenDone.toString());
						if (blnOpenWhenDone) {
							//**********	open the file in Excel
							Application.Shell(strExcel, "\"" + strFile + "\"");
						}
						
						
						sOut = strFile;
					}
					catch(e) {
						var errMessage = e.toString();
						switch (errMessage) {
							case "INVALID_DIRECTORY":
								Console.Writeln("The specified directory " + strDirectory + " could not be found. Export failed!", "Error exporting to Excel");
								break;
							case "NO_DIRECTORY":
								Console.Writeln("No directory was specified.", "Error exporting to Excel");
								break;
							case "NO_FILENAME":
								Console.Writeln("No file name was specified.", "Error exporting to Excel");
								break;
							case "BAD_DIRECTORY":
								Console.Writeln("The directory name must be a string data type.", "Error exporting to Excel");
								break;
							case "BAD_FILENAME":
								Console.Writeln("The file name must be a string data type.", "Error exporting to Excel");
								break;
							case "FILE_STILL_OPEN":
								Console.Writeln("\"" + strFile + "\" is still open or read only.", "Error exporting to Excel");
								break;
							case "UNKNOWN_OPTION":
								Console.Writeln("An unknown option was chosen.  This shouldn't be possible.  Please contact your administrator.", "Error exporting to Excel");
								break;
							default:
								Console.Writeln(errMessage, "Error exporting to Excel");
						}
						sOut = errMessage;
					}
					finally {
					}
					
					//Console.Writeln("Export complete");
					return sOut;
				}	//	Section.ExportToExcel
				
				if (oSec.Type == bqResults || oSec.Type == bqTable) {
					//Console.Writeln(oSec.Name);
					var oTable = oSec;
					oTable.sortList = new Array();		//	sample item value:  "columnname|order"
					oTable.sort = function (column, order) {
						//	column		string				name of column to sort
						//	order		bqSortOrder value	1 = ascending, 2 = descending
						//	If no arguments are passed, clear the sort items.
						//Console.Writeln(this.Name);
						//Console.Writeln(this.Type);
						
						var sMsg = "";
						if (typeof order == "undefined") order = 1;
						
						if (typeof column == "undefined") {
							//	clear sort list
							this.sortList.splice(0, this.sortList.length);
							this.SortItems.RemoveAll();
							this.SortItems.SortNow();
						}
						else {
							var blnValidInput = false;
							if (typeof column == "string") {
								try {
									var z = this.Columns.Item(column);	//	this verifies that column is a column
									if (typeof order == "number") {
										if (order == 1 || order == 2) {
											//	column is a string and the name of a column
											//	and order is a number and either 1 or 2
											blnValidInput = true;
										}
										else {
											sMsg += order + " must be 1 or 2.\r\n";
											blnValidInput = false;
										}
									}
									else {
										sMsg += order + " is not a number.\r\n";
										blnValidInput = false;
									}
								}
								catch (e) {
									//	column isn't a column
									sMsg += column + " is not the name of a column.\r\n";
									blnValidInput = false;
								}
							}
							else {
								sMsg += column + " is not a string.\r\n";
								blnValidInput = false;
							}
							
							if (blnValidInput) {
								//	it's all good, let's get to work
								//	if the column is already in the sort list, remove it
								this.sortList.remove(column + "|1");
								this.sortList.remove(column + "|2");
								//	add the column to the beginning of the sort list
								this.sortList.unshift(column + "|" + order.toString());
								
								//	since the names of the sort items are not available, this gets a little weird
								//	remove all sort items
								this.SortItems.RemoveAll();
								//	add the sort items in the correct order
								for (var i = 0; i < this.sortList.length; i++) {
									this.SortItems.Add(this.sortList[i].substr(0, this.sortList[i].indexOf("|")));
									this.SortItems.Item(i + 1).SortOrder = 1.0 * this.sortList[i].substr(this.sortList[i].length - 1);
								}
								//	tell Hyperion to sort the table
								this.SortItems.SortNow();
							}
							else {
								Console.Writeln(sMsg);
							}
						}
					};	//	sort
					oTable.sortAppend = function (column, order) {
						//	column		string				name of column to sort
						//	order		bqSortOrder value	1 = ascending, 2 = descending
						//	If no arguments are passed, clear the sort items.
						//Console.Writeln(this.Name);
						//Console.Writeln(this.Type);
						
						var sMsg = "";
						if (typeof order == "undefined") order = 1;
						
						if (typeof column == "undefined") {
							//	clear sort list
							this.sortList.splice(0, this.sortList.length);
							this.SortItems.RemoveAll();
							this.SortItems.SortNow();
						}
						else {
							var blnValidInput = false;
							if (typeof column == "string") {
								try {
									var z = this.Columns.Item(column);	//	this verifies that column is a column
									if (typeof order == "number") {
										if (order == 1 || order == 2) {
											//	column is a string and the name of a column
											//	and order is a number and either 1 or 2
											blnValidInput = true;
										}
										else {
											sMsg += order + " must be 1 or 2.\r\n";
											blnValidInput = false;
										}
									}
									else {
										sMsg += order + " is not a number.\r\n";
										blnValidInput = false;
									}
								}
								catch (e) {
									//	column isn't a column
									sMsg += column + " is not the name of a column.\r\n";
									blnValidInput = false;
								}
							}
							else {
								sMsg += column + " is not a string.\r\n";
								blnValidInput = false;
							}
							
							if (blnValidInput) {
								//	it's all good, let's get to work
								//	if the column is already in the sort list, remove it
								this.sortList.remove(column + "|1");
								this.sortList.remove(column + "|2");
								//	add the column to the beginning of the sort list
								this.sortList.push(column + "|" + order.toString());
								
								//	since the names of the sort items are not available, this gets a little weird
								//	remove all sort items
								this.SortItems.RemoveAll();
								//	add the sort items in the correct order
								for (var i = 0; i < this.sortList.length; i++) {
									this.SortItems.Add(this.sortList[i].substr(0, this.sortList[i].indexOf("|")));
									this.SortItems.Item(i + 1).SortOrder = 1.0 * this.sortList[i].substr(this.sortList[i].length - 1);
								}
								//	tell Hyperion to sort the table
								this.SortItems.SortNow();
							}
							else {
								Console.Writeln(sMsg);
							}
						}
					};	//	sortAppend
					oSec.getValues = function (column) {
						//	return an array of distinct values from the specified column
						var a = new Array();
						var tmpLim = this.Limits.CreateLimit(column);
						
						tmpLim.RefreshAvailableValues();
						
						for (i = 1; i <= tmpLim.AvailableValues.Count; i++) {
							a.push(tmpLim.AvailableValues[i]);
						}
						
						return a;
					};	//	getValues
					try { 
						for (var c = 1; c <= oSec.Columns.Count; c++) {
							oCol = oSec.Columns.Item(c);
							oCol.Parent = oSec;
							oCol.toArray = function() {
								//	extract the data to an array
								var arr = new Array();
								var i = 1;
								var val;
								try { val = this.GetCell(i++) }
								catch (e) {
									//	no records
									//throw new Error("no records");
									return arr;
								}
								
								for (var blnDone = false; !blnDone; ) {
									arr.push(val);
									try { val = this.GetCell(i++) }
									catch (e) { blnDone = true; }
								}
								////	eliminate duplicates and sort the array
								//arr = arr.uniq(true);
								return arr;
							};	//	Column.toArray
						};
					}
					catch (e) {
						//	no columns found
						Console.Writeln("No columns found:  " + oSec.Name);
						//	this may be a result from a stored procedure that doesn't return data
					}
				}
			}
			
			if (([bqReport]).contains(oSec.Type)) {
				oSec.ExportToExcel = function (strDirectory, strExpFileName, blnOpenWhenDone, blnUnmergeCells, blnUnPivot) {
					/***********************************************************************************
					*	Code copied from http://it.toolbox.com/wiki/index.php/Export_Pivots_and_Results_from_BQY_to_Excel_with_Style
					*	
					*	This function takes the following 4 arguments
					*	Arg Name			Arg Type	Arg Description
					*	strDirectory		String		The directory to export to
					*	strExpFileName		String		The name to use for the exported file
					*									(*exclude file extension)
					*	blnOpenWhenDone		Boolean		indicates whether to open the file after exporting
					*	blnUnmergeCells		Boolean		indicates whether to unmerge merged cells (when exporting from a pivot)
					*	blnUnPivot			Boolean		indicates whether to unpivot the data from a Pivot, putting the column labels before the row labels
					*	
					*	Return value:		the path and name of the file created
					*	
					*	Sample function call:
					*	ExportToExcel("C:\\temp", "Report0207-A", true);
					*	This would create a file named Report0207-A.xls in the temp directory on the 
					*	user's C drive containing the contents of the section object and open the file.
					***********************************************************************************/
					
					//	development note:
					//		use blnUnPivot instead of blnKeepFormatting
					//		Ask the user "how do you want it displayed", "as-is", "unmerged", "unpivotted"
					//		
					
					
					//Console.Writeln("Exporting section to " + strDirectory + "\\" + strExpFileName + ".xls");
					var sOut;
					
					try {
						//**********	validate input
						if (typeof strDirectory == "undefined" || strDirectory == "") throw("NO_DIRECTORY");
						if (typeof strDirectory != "string") throw("BAD_DIRECTORY");
						if (typeof strExpFileName == "undefined" || strExpFileName == "") throw("NO_FILENAME");
						if (typeof strExpFileName != "string") throw("BAD_FILENAME");
						if (typeof blnOpenWhenDone == "undefined") blnOpenWhenDone = false;
						blnOpenWhenDone = new Boolean(blnOpenWhenDone);
						
						if (this.Type == bqPivot) {
							//	If it's not a Pivot, we don't care about the last 2 parameters.
							if (typeof blnUnmergeCells == "undefined" && typeof blnUnPivot == "undefined") {
								var iHow = Alert("How do you want the data exported?", "Export To Excel", "Merged", "Unmerged", "Unpivotted");
								switch (iHow) {
									case 1:
										blnUnmergeCells = false;
										blnUnPivot = false;
										break;
									case 2:
										blnUnmergeCells = true;
										blnUnPivot = false;
										break;
									case 3:
										blnUnmergeCells = false;
										blnUnPivot = true;
										break;
									default:
										throw("UNKNOWN_OPTION");
										break;
								}
							}
							else if (typeof blnUnmergeCells == "undefined" && !blnUnPivot) {
								//	blnUnPivot is defined
								blnUnPivot = new Boolean(blnUnPivot);
								//	If the data is unpivotted, it's already unmerged.
								//	if blnUnPivot is false, ask about unmerging
								if (!blnUnPivot) {
									var iUnMerge = 2;
									
									if (this.Type == bqPivot) {
										iUnMerge = Alert("Do you want the cells unmerged?", "Export to Excel", "Yes", "No");
									}
									
									if (iUnMerge == 1) {
										blnUnmergeCells = true;
									}
									else {
										blnUnmergeCells = false;
									}
								}
							}
							else if (typeof blnUnPivot == "undefined") {
								//	blnUnmergeCells is defined
								blnUnmergeCells = new Boolean(blnUnmergeCells);
								//	If the data is unmerged, it can still be pivotted.
								//	if blnUnmergeCells is true, ask about unpivotting
								if (blnUnmergeCells) {
									var iUnPivot = 2;
									
									if (this.Type == bqPivot) {
										iUnPivot = Alert("Do you want the cells unpivotted?", "Export to Excel", "Yes", "No");
									}
									
									if (iUnPivot == 1) {
										blnUnPivot = true;
									}
									else {
										blnUnPivot = false;
									}
								}
							}
						}
						else {
							blnUnmergeCells = false;
							blnUnPivot = false;
						}
						
						if (blnUnmergeCells) {
							strFileExt = ".xls";
						}
						else {
							strFileExt = ".mhtml";
						}
						
						//**********	Verify The Directory Exists
						if (!Application.FSO.FolderExists(strDirectory)) {
							throw("INVALID_DIRECTORY");
						}
						//**********	ensure the directory name ends in a slash
						if (strDirectory.right(1) != "\\") strDirectory += "\\";
						//Console.Writeln("	folder exists");
						
						var strFile = strDirectory + strExpFileName + strFileExt;
						//Console.Writeln("	file names defined");
						
						//**********	Turn Off Page Breaks
						this.HTMLHorizontalPageBreakEnabled = false;
						this.HTMLVerticalPageBreakEnabled = false;
						
						//**********	If Files Already Exist - Delete Them
						if (Application.FSO.FileExists(strFile)) {
							try {
								Application.FSO.DeleteFile (strFile);
							}
							catch(e) {
								if (e.toString() != "") {
									throw("FILE_STILL_OPEN");
								}
							}
						}
						
						//Console.Writeln("	ready to export");
						
						if (blnUnPivot) {
							//Console.Writeln("keep formatting");
							//	create a table based on the pivot and export it
							var arrET = this.ExportTable();
							arrET[0].ExportToStream(strFile, bqExportFormatOfficeMHTML);
							for (var i = 0; i < arrET.length; i++) {
								arrET[i].Remove();
							}
						}
						else if (blnUnmergeCells) {
							//Console.Writeln("unmerge");
							//	export the Pivot as data for further analysis
							//	remove totals before exporting
							var arrTop = new Array();
							var arrSide = new Array();
							var oTotal;
							
							for (var i = 1; i <= this.TopLabels.Count; i++) {
								for (var j = 1; j = this.TopLabels.Item(i).Totals.Count; j++) {
									arrTop.push([this.TopLabels.Item(i).Name, this.TopLabels.Item(i).Totals.Item(j).DataFunction]);
									this.TopLabels.Item(i).Totals.Item(j).Remove();
								}
							}
							
							for (var i = 1; i <= this.SideLabels.Count; i++) {
								for (var j = 1; j = this.SideLabels.Item(i).Totals.Count; j++) {
									arrSide.push([this.SideLabels.Item(i).Name, this.SideLabels.Item(i).Totals.Item(j).DataFunction]);
									this.SideLabels.Item(i).Totals.Item(j).Remove();
								}
							}
							
							//**********	Export to Excel 97/2003 Format
							this.ExportToStream(strFile, bqExportFormatExcel8);
							//Console.Writeln("	xls file exported");
							
							//	add the totals that were removed, in reverse order to match what was there
							for (j = arrSide.length - 1; j >= 0; j--) {
								for (var i = 1; i <= this.SideLabels.Count; i++) {
									if (arrSide[j][0] == this.SideLabels.Item(i).Name) {
										oTotal = this.SideLabels.Item(i).Totals.Add();
										oTotal.DataFunction = arrSide[j][1];
									}
								}
							}
							
							for (j = arrTop.length - 1; j >= 0; j--) {
								for (var i = 1; i <= this.TopLabels.Count; i++) {
									if (arrTop[j][0] == this.TopLabels.Item(i).Name) {
										oTotal = this.TopLabels.Item(i).Totals.Add();
										oTotal.DataFunction = arrTop[j][1];
									}
								}
							}
						}
						else {
							//**********	Export to MHTML Format
							this.ExportToStream(strFile, bqExportFormatOfficeMHTML);
							//Console.Writeln("	mhtml file exported");
						}
						
						////**********	Get path to EXCEL executable
						//var objExcel = new JOOLEObject("Excel.Application");
						//Console.Writeln("	created Excel object");
						//var strExcel = objExcel.Path + "\\Excel.exe";
						//objExcel.DisplayAlerts = false;
						//objExcel.Quit();
						var strExcel = (new JOOLEObject("Excel.Application")).Path + "\\Excel.exe";
						
						//Console.Writeln(blnOpenWhenDone.toString());
						if (blnOpenWhenDone) {
							//**********	open the file in Excel
							Application.Shell(strExcel, "\"" + strFile + "\"");
						}
						
						
						sOut = strFile;
					}
					catch(e) {
						var errMessage = e.toString();
						switch (errMessage) {
							case "INVALID_DIRECTORY":
								Console.Writeln("The specified directory " + strDirectory + " could not be found. Export failed!", "Error exporting to Excel");
								break;
							case "NO_DIRECTORY":
								Console.Writeln("No directory was specified.", "Error exporting to Excel");
								break;
							case "NO_FILENAME":
								Console.Writeln("No file name was specified.", "Error exporting to Excel");
								break;
							case "BAD_DIRECTORY":
								Console.Writeln("The directory name must be a string data type.", "Error exporting to Excel");
								break;
							case "BAD_FILENAME":
								Console.Writeln("The file name must be a string data type.", "Error exporting to Excel");
								break;
							case "FILE_STILL_OPEN":
								Console.Writeln("\"" + strFile + "\" is still open or read only.", "Error exporting to Excel");
								break;
							case "UNKNOWN_OPTION":
								Console.Writeln("An unknown option was chosen.  This shouldn't be possible.  Please contact your administrator.", "Error exporting to Excel");
								break;
							default:
								Console.Writeln(errMessage, "Error exporting to Excel");
						}
						sOut = errMessage;
					}
					finally {
					}
					
					//Console.Writeln("Export complete");
					return sOut;
				},	//	Section.ExportToExcel
				
				oSec.ExportToPDF = function (strDirectory, strExpFileName, blnOpenWhenDone) {
					/***********************************************************************************
					*	This function takes the following 4 arguments
					*	Arg Name			Arg Type	Arg Description
					*	strDirectory		String		The directory to export to
					*	strExpFileName		String		The name to use for the exported file
					*									(*exclude file extension)
					*	blnOpenWhenDone		Boolean		indicates whether to open the file after exporting
					*	
					*	Return value:		the path and name of the file created
					*	
					*	Sample function call:
					*	ExportToPDF("C:\\temp", "Report0207-A", true);
					*	This would create a file named Report0207-A.pdf in the temp directory on the 
					*	user's C drive containing the contents of the section object and open the file.
					***********************************************************************************/
					
					var sOut;
					
					try {
						//**********	validate input
						if (typeof strDirectory == "undefined" || strDirectory == "") throw("NO_DIRECTORY");
						if (typeof strDirectory != "string") throw("BAD_DIRECTORY");
						if (typeof strExpFileName == "undefined" || strExpFileName == "") throw("NO_FILENAME");
						if (typeof strExpFileName != "string") throw("BAD_FILENAME");
						if (typeof blnOpenWhenDone == "undefined") blnOpenWhenDone = false;
						blnOpenWhenDone = new Boolean(blnOpenWhenDone);
						
						strFileExt = ".pdf";
						
						//**********	Verify The Directory Exists
						if (!Application.FSO.FolderExists(strDirectory)) {
							throw("INVALID_DIRECTORY");
						}
						//**********	ensure the directory name ends in a slash
						if (strDirectory.right(1) != "\\") strDirectory += "\\";
						
						var strFile = strDirectory + strExpFileName + strFileExt;
						
						//**********	If Files Already Exist - Delete Them
						if (Application.FSO.FileExists(strFile)) {
							try {
								Application.FSO.DeleteFile (strFile);
							}
							catch(e) {
								if (e.toString() != "") {
									throw("FILE_STILL_OPEN");
								}
							}
						}
						
						//**********	Export to pdf Format
						this.ExportToStream(strFile, bqExportFormatPDF);
						
						if (blnOpenWhenDone) {
							//**********	open the file in default program for PDF files
							Application.Shell("explorer.exe", "\"" + strFile + "\"");
						}
						
						sOut = strFile;
					}
					catch(e) {
						var errMessage = e.toString();
						switch (errMessage) {
							case "INVALID_DIRECTORY":
								Console.Writeln("The specified directory " + strDirectory + " could not be found. Export failed!", "Error exporting to PDF");
								break;
							case "NO_DIRECTORY":
								Console.Writeln("No directory was specified.", "Error exporting to PDF");
								break;
							case "NO_FILENAME":
								Console.Writeln("No file name was specified.", "Error exporting to PDF");
								break;
							case "BAD_DIRECTORY":
								Console.Writeln("The directory name must be a string data type.", "Error exporting to PDF");
								break;
							case "BAD_FILENAME":
								Console.Writeln("The file name must be a string data type.", "Error exporting to PDF");
								break;
							case "FILE_STILL_OPEN":
								Console.Writeln("\"" + strFile + "\" is still open or read only.", "Error exporting to PDF");
								break;
							case "UNKNOWN_OPTION":
								Console.Writeln("An unknown option was chosen.  This shouldn't be possible.  Please contact your administrator.", "Error exporting to PDF");
								break;
							default:
								Console.Writeln(errMessage, "Error exporting to PDF");
						}
						sOut = errMessage;
					}
					finally {
					}
					
					return sOut;
				}	//	Section.ExportToPDF
			}
			
			if (oSec.Type == bqPivot) {
				oSec.ExportTable = function () {
					//	this.Parent is the Sections collection containing the Pivot
					//	this.Parent.Parent is the Document containing the Pivot
					var arrTop = new Array();
					var arrSide = new Array();
					var arrFact = new Array();
					var arrFactC = new Array();
					var sKey = "";
					var sName = this.ParentSection.Name;
					
					for (var i = 1; i <= this.TopLabels.Count; i++) {
						arrTop.push(this.TopLabels.Item(i).Name);
						//	IR converts !@#%^&^*()_~+-=`\|[{]},./<>?;':" to underscores
						//sKey += " + " + this.TopLabels.Item(i).Name.toUnderscoreCase();
						sKey += " + " + this.TopLabels.Item(i).Name.replace(/[ `~!@#%\^&\*\(\)_\+\-=\[\]\\\{\}\|;':",\.\/<>\?]/gi, "_");
					}
					
					for (var i = 1; i <= this.SideLabels.Count; i++) {
						arrSide.push(this.SideLabels.Item(i).Name);
						//	IR converts !@#%^&^*()_~+-=`\|[{]},./<>?;':" to underscores
						//sKey += " + " + this.SideLabels.Item(i).Name.toUnderscoreCase();
						sKey += " + " + this.SideLabels.Item(i).Name.replace(/[ `~!@#%\^&\*\(\)_\+\-=\[\]\\\{\}\|;':",\.\/<>\?]/gi, "_");
					}
					
					for (var i = 1; i <= this.Facts.Count; i++) {
						arrFact.push(this.Facts.Item(i).Name);
					}
					
					sKey = sKey.substr(3);
					
					//-------------------------------------------------------------------------
					//	create and configure interim table
					var tInterim = this.Parent.Add(bqTable, sName);
					tInterim.Name = "T-" + sName + " interim";
					
					for (var i = 0; i < arrTop.length; i++) {
						tInterim.Columns.Add(arrTop[i]);
						tInterim.SortItems.Add(arrTop[i]);
					}
					
					for (var i = 0; i < arrSide.length; i++) {
						tInterim.Columns.Add(arrSide[i]);
						tInterim.SortItems.Add(arrSide[i]);
					}
					
					for (var i = 0; i < arrFact.length; i++) {
						tInterim.Columns.Add(arrFact[i]);
					}
					
					tInterim.Columns.AddComputed("key", sKey);
					
					for (var i = 0; i < arrFact.length; i++) {
						tInterim.Columns.AddComputed(" " + arrFact[i] + " ", "Sum ( " + arrFact[i].toUnderscoreCase() + ", key )");
						arrFactC.push(" " + arrFact[i] + " ");
					}
					
					tInterim.SortItems.SortNow();
					tInterim.Columns.AddComputed("filter", "if ( Prior ( key ) != key ) { 1 }");
					
					
					
					//-------------------------------------------------------------------------
					//	create and configure filtered table
					var tFilter = this.Parent.Add(bqTable, tInterim.Name);
					tFilter.Name = "T-" + sName + " filter";
					
					for (var i = 0; i < arrTop.length; i++) {
						tFilter.Columns.Add(arrTop[i]);
						tFilter.SortItems.Add(arrTop[i]);
					}
					
					for (var i = 0; i < arrSide.length; i++) {
						tFilter.Columns.Add(arrSide[i]);
						tFilter.SortItems.Add(arrSide[i]);
					}
					
					for (var i = 0; i < arrFactC.length; i++) {
						tFilter.Columns.Add(arrFactC[i]);
					}
					
					tFilter.SortItems.SortNow();
					tFilter.Columns.Add("filter");
					var oLim = tFilter.Limits.CreateLimit("filter");
					oLim.SelectedValues.Add(1);
					tFilter.Limits.Add(oLim);
					
					
					//-------------------------------------------------------------------------
					//	create and configure output table
					var tOutput = this.Parent.Add(bqTable, tFilter.Name);
					tOutput.Name = "T-" + sName + " output";
					
					for (var i = 0; i < arrTop.length; i++) {
						tOutput.Columns.Add(arrTop[i]);
						tOutput.SortItems.Add(arrTop[i]);
					}
					
					for (var i = 0; i < arrSide.length; i++) {
						tOutput.Columns.Add(arrSide[i]);
						tOutput.SortItems.Add(arrSide[i]);
					}
					
					for (var i = 0; i < arrFactC.length; i++) {
						tOutput.Columns.Add(arrFactC[i]);
					}
					
					tOutput.SortItems.SortNow();
					
					this.Parent.Parent.loadObjectMethods();
					
					return [tOutput, tFilter, tInterim];
				}	//	ExportTable
			}
			
			if ((new Array(bqResults, bqTable, bqQuery)).contains(oSec.Type)) {
				for (var k = 1; k <= oSec.Limits.Count; k++) {
					oSec.Limits[k].getSelection = function () {
						var a = new Array();
						for (var j = 1; j <= this.SelectedValues.Count; j++) {
							a.push(this.SelectedValues[j]);
						}
						return a;
					}
					oSec.Limits[k].setSelection = function (arr) {
						var a = new Array();
						if (arr.toString() == "[object JBLimit]") {
							a = arr.getSelection();
						}
						else if (Array.prototype.isPrototypeOf(arr)) {
							a = arr.uniq(true);		//	changed 9/28/2012
						}
						this.SelectedValues.RemoveAll();
						for (var j = 0; j < a.length; j++) {
							this.SelectedValues.Add(a[j]);
						}
					}
				}
				//	added these methods to Limits in AppendQueries
				//	dp - 10/22/2012
				if (oSec.Type == bqQuery) {
					for (var m = 1; m <= oSec.AppendQueries.Count; m++) {
						for (var k = 1; k <= oSec.AppendQueries.Item(m).Limits.Count; k++) {
							oSec.AppendQueries.Item(m).Limits[k].getSelection = function () {
								var a = new Array();
								for (var j = 1; j <= this.SelectedValues.Count; j++) {
									a.push(this.SelectedValues[j]);
								}
								return a;
							}
							oSec.AppendQueries.Item(m).Limits[k].setSelection = function (arr) {
								var a = new Array();
								if (arr.toString() == "[object JBLimit]") {
									a = arr.getSelection();
								}
								else if (Array.prototype.isPrototypeOf(arr)) {
									a = arr.uniq(true);
								}
								this.SelectedValues.RemoveAll();
								for (var j = 0; j < a.length; j++) {
									this.SelectedValues.Add(a[j]);
								}
							}
						}
					}
				}
			}
		}
	}, 	//	loadObjectMethods
	
	ProgressBar: Class.create({
		//	currently this is just a blue bar in a black box
		//	Additional appearance proposals:
		//		sunken box
		//		adjustable border color
		//		adjustable bar color
		//		gradient bar color
		
		initialize: function(dashboard, left, top, width, height, value, maxvalue, label, font) {
			//	left		number	location of left of progressbar
			//	top			number	location of top of progressbar
			//	width		number	width of progressbar
			//	height		number	height of progressbar
			//	value		number	current value to display
			//	maxvalue	number	maximum value to display
			//	label		boolean	Should a label be incorporated?
			
			_left = left == undefined ? 0 : left;
			_left = typeof _left == "number" ? _left : 0;
			_top = top == undefined ? 0 : top;
			_top = typeof _top == "number" ? _top : 0;
			_width = width == undefined ? 100 : width;
			_width = typeof _width == "number" ? _width : 100;
			_height = height == undefined ? 20 : height;
			_height = typeof _height == "number" ? _height : 20;
			_val = value == undefined ? 0 : value;
			_val = typeof _val == "number" ? _val : 0;
			_maxval = maxvalue == undefined ? 100 : maxvalue;
			_maxval = typeof _maxval == "number" ? _maxval : 100;
			if (_maxval == 0) return "maxvalue can't be 0";
			_label = label == undefined ? false : label;
			_label = typeof _label == "boolean" ? _label : false;
			_font = font == undefined ? new Application.Font(undefined, undefined, new Application.Font().StyleBold(), undefined, "00cc00") : font;
			_font = _font.isFont() ? _font : new Application.Font(undefined, undefined, new Application.Font().StyleBold(), undefined, "00cc00");
			_curr = _width * _val / _maxval;
			_name = "";
			
			//	hyperion stuff
			if (dashboard == undefined) return "Dashboard required";
			try { blnTemp = dashboard.Type == bqDashboard; }
			catch (e) { return "The reference does not refer to an object."; }
			if (!blnTemp) return "The reference does not refer to a dashboard.";
			_dash = dashboard;
			_rctOuter = _dash.Shapes.CreateShape(bqRectangle);
			with (_rctOuter) {
				Name = "rctProg1_" + (new Date()).valueOf().toString();
				Visible = false;
				Placement.Modify(_left, _top, _width, _height);
				Line.Width = 1;
				Line.Color = 0;
				Line.DashStyle = bqDashStyleSolid;
			}
			_rctInner = _dash.Shapes.CreateShape(bqRectangle);
			with (_rctInner) {
				Name = "rctProg2_" + (new Date()).valueOf().toString();
				Visible = false;
				Placement.Modify(_left, _top, _curr, _height);
				Line.Width = 0;
				Fill.Pattern = bqFillPattern100;
				Fill.Color = parseInt("0000ff", 16);
			}
			_lblPB = _dash.Shapes.CreateShape(bqTextLabel);
			with (_lblPB) {
				Name = "lblProg_" + (new Date()).valueOf().toString();
				Visible = false;
				Placement.Modify(_left, _top, _width, _height);
				Line.Width = 0;
				Fill.Pattern = bqFillPattern100;
				Fill.Color = parseInt("fefefe", 16);	//	transparent
				VerticalAlignment = 2;
				Alignment = 1;
				this.Font(_font);
			}
			if (_label) {
				_lblPB.Visible = true;
			}
			else {
				_lblPB.Visible = false;
			}
			//Console.Writeln(_dash.Name);
		},	//	initialize
			
		Left: function (vLeft) {
			_left = typeof vLeft == "number" ? vLeft : _left;
			_rctInner.Placement.XOffset = _left;
			_rctOuter.Placement.XOffset = _left;
			_lblPB.Placement.XOffset = _left;
			return _left;
		},	//	Left
		
		Top: function (vTop) {
			_top = typeof vTop == "number" ? vTop : _top;
			_rctInner.Placement.YOffset = _top;
			_rctOuter.Placement.YOffset = _top;
			_lblPB.Placement.YOffset = _top;
			return _top;
		},	//	Top
		
		Width: function (vWidth) {
			_width = typeof vWidth == "number" ? vWidth : _width;
			_curr = _width * _val / _maxval;
			_rctInner.Placement.Width = _curr;
			_rctOuter.Placement.Width = _width;
			_lblPB.Placement.Width = _width;
			return _width;
		},	//	Width
		
		Height: function (vHeight) {
			_height = typeof vHeight == "number" ? vHeight : _height;
			_rctInner.Placement.Height = _height;
			_rctOuter.Placement.Height = _height;
			_lblPB.Placement.Height = _height;
			return _height;
		},	//	Height
		
		Value: function (vVal) {
			_val = typeof vVal == "number" ? vVal : _val;
			_val = _val < _maxval ? _val : _maxval;
			_curr = _width * _val / _maxval;
			//	hyperion stuff
			_rctInner.Placement.Width = _curr;
			//_rctOuter.Placement.Width = _width;
			_rctOuter.Fill.Pattern = bqFillPattern100;
			return _val;
		},	//	Value
		
		MaxValue: function (vMaxVal) {
			_maxval = typeof vMaxVal == "number" ? vMaxVal : _maxval;
			_curr = _width * _val / _maxval;
			_rctInner.Placement.Width = _curr;
			//_rctOuter.Placement.Width = _width;
			return _maxval;
		},	//	MaxValue
		
		Font: function (vFont) {
			_font = vFont == undefined ? _font : (vFont.isFont() ? vFont : _font);
			_lblPB.Font.Color = _font.Color();
			_lblPB.Font.Effect = _font.Effect();
			_lblPB.Font.Name = _font.Name();
			_lblPB.Font.Size = _font.Size();
			_lblPB.Font.Style = _font.Style();
			return _font;
		},	//	Font
		
		Label: function (vLabel) {
			_label = typeof vLabel == "boolean" ? vLabel : _label;
			//	hyperion stuff
			if (_rctOuter.Visible) _lblPB.Visible = true;
			return _label;
		},	//	Label
		
		Text: function (vText) {
			_text = typeof vText == "string" ? vText : _text;
			_lblPB.Text = _text;
			return _text;
		},	//	Text
		
		Show: function () {
			_rctInner.Visible = true;
			_rctOuter.Visible = true;
			if (_label) _lblPB.Visible = true;
		},	//	Show
		
		Hide: function () {
			_rctInner.Visible = false;
			_rctOuter.Visible = false;
			_lblPB.Visible = false;
		},	//	Hide
		
		Kill: function () {
			ActiveDocument.Sections[_rctInner.Parent.Name].Shapes.RemoveShape(_rctInner.Name);
			ActiveDocument.Sections[_rctOuter.Parent.Name].Shapes.RemoveShape(_rctOuter.Name);
			ActiveDocument.Sections[_lblPB.Parent.Name].Shapes.RemoveShape(_lblPB.Name);
		}	//	Kill
	}), 	//	ProgressBar
	
	PartitionFilter: Class.create({
		//	This can be further automated by defining all of the partitioning 
		//	information for each data mart
		//	It may be possible to automatically generate the objects (like 
		//	_qryMetadata and _qryPartData) on the fly so the user doesn't have 
		//	to manually create them.
		//initialize: function(oQueryToFilter, oFilterItem, oPartitionValuesResult, oPartitionValuesColumn, oMetadataResult) {
		initialize: function(oQueryToFilter, oFilterItem, oPartitionValuesResult, oPartitionValuesColumn) {
			_limFactTableName = undefined;
			_qryToFilter = undefined;
			_sFactTable = undefined;
			_limPartition = undefined;
			_qryPartData = undefined;
			_rstPartData = undefined;
			_colPartData = undefined;
			//_rstMetadata = undefined;
			//_qryMetadata = undefined;
			
			this.QueryToFilter(oQueryToFilter);
			if (typeof _qryToFilter != "undefined") {
				//	if _qryToFilter is undefined, there's no point in looking for stuff in it
				if (typeof oFilterItem != "undefined") {
					this.FilterItem(oFilterItem);
				}
			}
			
			if (typeof oPartitionValuesResult != "undefined") {
				this.PartitionValuesResult(oPartitionValuesResult);
			}
			else {
				if (typeof _rstPartData == "undefined") {
					try {
						//	check for the existence of a result named r-PartData
						this.PartitionValuesResult(ActiveDocument.Sections["r-PartData"]);
					}
					catch (e) {
						//	result not found
					}
				}
			}
			
			if (typeof oPartitionValuesColumn != "undefined") {
				this.PartitionValuesColumn(oPartitionValuesColumn);
			}
			
			//if (typeof oMetadataResult != "undefined") {
			//	this.MetadataResult(oMetadataResult);
			//}
			//else {
			//	if (typeof _rstMetadata == "undefined") {
			//		try {
			//			//	check for the existence of a result named r-Partition
			//			this.MetadataResult(ActiveDocument.Sections["r-Partition"]);
			//		}
			//		catch (e) {
			//			//	query not found
			//		}
			//	}
			//}
			
		},
		
		QueryToFilter: function (o) {
			//	Is this a Query object?
			if (typeof o == "object" && o.Type == bqQuery) {
				_qryToFilter = o;
				if (typeof _limPartition == "undefined") {
					//	_limPartition is undefined
					//	check for the existence of a limit named PartitionId in _qryToFilter
					try {
						this.FilterItem(_qryToFilter.Limits["PartitionId"]);
					}
					catch (e) {
						//	limit not found
					}
				}
				else {
					//	_limPartition is not undefined
					//	_limPartition must be in QueryToFilter
					this.FilterItem(_limPartition);
				}
			}
			return _qryToFilter;
		}, 
		
		FilterItem: function (o) {
			//	Is this a Limit object?
			//	Is _qryToFilter defined?
			if (typeof o == "object" && o.toString() == "[object JBLimit]" && typeof _qryToFilter != "undefined") {
				//	Is this in _qryToFilter?
				try {
					vTest = _qryToFilter.Limits[o.Name].CustomValues.Count;
					//	Is this in _sFactTable?
					_sFactTable = undefined;
					sTopic = o.FullName.substr(0, o.FullName.lastIndexOf("."));
					
					for (var i = 1; i <= _qryToFilter.DataModel.Topics.Count; i++) {
						if (_qryToFilter.DataModel.Topics[i].Name.toUnderscoreCase() == sTopic) {
							_sFactTable = _qryToFilter.DataModel.Topics[i].Name;
							break;
						}
					}
					
					if (typeof _sFactTable != "undefined") {
						try {
							vTest = _qryToFilter.DataModel.Topics[_sFactTable].TopicItems[o.FullName.substr(o.FullName.lastIndexOf(".") + 1)];
							_limPartition = o;
						}
						catch (e) {
							//	o isn't for an item in _sFactTable
							_limPartition = undefined;
						}
					}
					else {
						//	o's Topic wasn't found in the data model
						//	this shouldn't happen
						_limPartition = undefined;
					}
				}
				catch (e) {
					//	o doesn't exist or isn't a limit object in _qryToFilter
					_limPartition = undefined;
				}
			}
			return _limPartition;
		}, 
		
		//MetadataResult: function (o) {
		//	if (typeof o == "object" && o.Type == bqResults) {
		//		try {
		//			_limFactTableName = o.Limits["Fact Table Name"];
		//			_rstMetadata = o;
		//			_qryMetadata = o.ParentSection;
		//		}
		//		catch (e) {
		//			_limFactTableName = undefined;
		//			_rstMetadata = undefined;
		//			_qryMetadata = undefined;
		//		}
		//	}
		//	return _rstMetadata;
		//}, 
		
		PartitionValuesResult: function (o) {
			if (typeof o == "object" && o.Type == bqResults) {
				_rstPartData = o;
				_qryPartData = o.ParentSection;
				if (typeof _colPartData != "undefined") {
					//	_colPartData is defined
					//	verify it's still good
					this.PartitionValuesColumn(_colPartData);
				}
				if (typeof _colPartData == "undefined") {
					//	_colPartData is undefined
					//	If _rstPartData has only one column, use it
					if (_rstPartData.Columns.Count == 1) {
						this.PartitionValuesColumn(_rstPartData.Columns[1]);
					}
					else {
						try {
							//	check for the existence of a column named PartitionId
							this.PartitionValuesColumn(_rstPartData.Columns["PartitionId"]);
						}
						catch (e) {
							//	column not found
						}
					}
				}
			}
			return _rstPartData;
		}, 
		
		PartitionValuesColumn: function (o) {
			//	Is o a column object?
			if (typeof o == "object" && o.toString() == "[object JBColumn]") {
				//	Is o in _rstPartData?
				if (o.Parent.Parent == _rstPartData) {
					//	it is
					_colPartData = o;
				}
				else {
					_colPartData = undefined;
				}
			}
			return _colPartData;
		}, 
		
		SetPartition: function () {
			//	variables used:
			//		_limFactTableName	Limit object in _rstMetadata to filter by Fact Table Name
			//		_qryToFilter		Query object to filter
			//		_sFactTable			Name Topic object (fact table) in _qryToFilter to filter
			//		_limPartition		Limit object in _qryToFilter to use to filter by partition
			//		_qryPartData		Query object to use to identify partition limit values
			//		_rstPartData		Result object to use to identify partition limit values
			//		_colPartData		Column object that contains the partition limit values
									
			if (typeof _qryToFilter == "undefined" || 
				typeof _limPartition == "undefined" || 
				typeof _qryPartData == "undefined" || 
				typeof _rstPartData == "undefined" || 
				typeof _colPartData == "undefined") return "error";
			//	typeof _limFactTableName == "undefined" || 
			//	typeof _sFactTable == "undefined" || 
				
			var blnUse = false;
			var sOut = "";
			var arrFilters = new Array();
			
			//_limFactTableName.SelectedValues.RemoveAll();
			//_limFactTableName.CustomValues.RemoveAll();
			//_limFactTableName.SelectedValues.Add(_sFactTable);
			
			//	ignore the partitioning limit in the query being processed
			//	in case we don't end up using it
			_limPartition.Ignore = true;
			
			//	ignore all filters in the partitioning query
			for (var i = 1; i <= _qryPartData.Limits.Count; i++) {
				_qryPartData.Limits.Item(i).Ignore = true;
				arrFilters.push(_qryPartData.Limits.Item(i).Name);
			}
			
			//	inspect filters in running query and use their values if possible
			for (var i = 1; i <= _qryToFilter.Limits.Count; i++) {
				//	don't bother if the filter is ignored
				if (!_qryToFilter.Limits.Item(i).Ignore) {
					//Console.Writeln("    " + _qryToFilter.Limits.Item(i).Name);
					//	is the filter one we can use?
					if (arrFilters.contains(_qryToFilter.Limits.Item(i).Name)) {
						//Console.Writeln("        " + _qryToFilter.Limits.Item(i).Name);
						blnUse = true;
						//	add the items to the partition filter
						_qryPartData.Limits[_qryToFilter.Limits.Item(i).Name].CustomValues.RemoveAll();
						_qryPartData.Limits[_qryToFilter.Limits.Item(i).Name].SelectedValues.RemoveAll();
						for (var j = 1; j <= _qryToFilter.Limits.Item(i).SelectedValues.Count; j++) {
							_qryPartData.Limits[_qryToFilter.Limits.Item(i).Name].CustomValues.Add(_qryToFilter.Limits.Item(i).SelectedValues.Item(j));
							_qryPartData.Limits[_qryToFilter.Limits.Item(i).Name].SelectedValues.Add(_qryToFilter.Limits.Item(i).SelectedValues.Item(j));
						}
						_qryPartData.Limits[_qryToFilter.Limits.Item(i).Name].Ignore = false;
					}
				}
			}
			
			if (blnUse) {
				//	we've determined that partitions may be useful
				_qryPartData.Process();
				_limPartition.CustomValues.RemoveAll();
				_limPartition.SelectedValues.RemoveAll();
				var n = _rstPartData.RowCount;
				if (n == 0) {
					_limPartition.CustomValues.Add(-1);
					_limPartition.SelectedValues.Add(-1);
				}
				else {
					//_limPartition.setSelection(_colPartData.toArray());
					_limPartition.setSelection(_colPartData.toArray().uniq(true));
				}
				_limPartition.Ignore = false;
				sOut = _limPartition.getSelection();
			}
			else {
				sOut = "not used";
			}
			
			return sOut;
		}, 
		
		toString: function() {
			return "[Object][PartitionFilter]";
		}
	})	//	PartitionFilter		//	added 9/28/2012
});


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
		
		//Day: function () {
		//	return _Day;
		//},
		//
		//DOW: function () {
		//	return _DOW;
		//},
		//
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
	}), 
	
	RadioButtonGroup: Class.create({
		Name: "", 
		Count: 0, 
		Type:  "RadioButtonGroup", 
		Parent: null, 
		
		initialize:  function(oParent, sGroupName) {
			_shapes = new Array();
			//var _parent = new Object();
			_groupName = ""
			//	oParent		required	dashboard object
			if (typeof oParent == "object") {
				if (oParent.Type == bqDashboard) {
					//	OK to continue
					this.Parent = oParent;
				}
				else {
					Console.Writeln("RadioButtonGroup:  '" + oParent.Name + "' is not a dashboard.");
					return;
				}
			}
			else {
				Console.Writeln("RadioButtonGroup:  Object not found.");
				return;
			}
			
			//	sGroupName	required	string
			if (typeof sGroupName == "string") {
				////	 sGroupName must not contain only letters and numbers
				//if (sGroupName.match(/\W/gi)) {
				//	Console.Writeln("RadioButtonGroup:  Invalid group name.");
				//	_groupName = "";
				//	return;
				//}
				//else {
					_groupName = sGroupName;
					this.Name = _groupName;
				//}
			}
			else {
				Console.Writeln("RadioButtonGroup:  GroupName must be a string.");
				_groupName = "";
				return;
			}
			
			//	capture all of the shapes that are already in the group
			for (var i = 1; i < this.Parent.Shapes.Count; i++) {
				if (this.Parent.Shapes.Item(i).Type == bqRadioButton) {
					if (this.Parent.Shapes.Item(i).Group == _groupName) {
						//	this is a group member
						this.Add(this.Parent.Shapes.Item(i));
					}
				}
			}
			
		}, 
		
		Add:  function (oShape) {
			//	shape must be a radio button object
			if (typeof oShape == "object") {
				if (oShape.Type == bqRadioButton) {
					if (oShape.Group != "") {
						//	the button a member of another group
						//	remove it from that group before adding it to this one
						for (var a in this.Parent) {
							if (typeof this.Parent[a] == "object") {
								try {
									if (this.Parent[a].Type == "RadioButtonGroup") {
										//	it's a RadioButtonGroup
										if (this.Parent[a].Members().contains(oShape)) {
											//	this shape is a member of the other RadioButtonGroup
											//	remove it
											this.Parent[a].Remove(oShape);
										}
									}
								}
								catch (e) {
								}
							}
						}
					}
					_shapes.push(oShape);
					oShape.Group = _groupName;
				}
			}
			this.Count = _shapes.length;
		}, 
		
		Remove:  function (oShape) {
			//	shape must be a radio button object
			if (typeof oShape == "object") {
				if (oShape.Type == bqRadioButton) {
					_shapes.remove(oShape);
					oShape.Group = "";
				}
			}
			this.Count = _shapes.length;
		}, 
		
		Members:  function () {
			return _shapes;
		}, 
		
		Value:  function () {
			var val = "";
			for (var i = 1; i < this.Parent.Shapes.Count; i++) {
				btn = this.Parent.Shapes.Item(i);
				if (btn.Type == bqRadioButton) {
					if (btn.Group == _groupName) {
						//	this is a group member
						//	is it selected?
						if (btn.Checked) {
							val = btn.Text;
						}
					}
				}
			}
			return val;
		}, 
		
		toString:  function () {
			return "[Object][RadioButtonGroup]";
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
