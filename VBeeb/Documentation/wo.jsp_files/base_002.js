function setImageURL() {}
function lib_bwcheck(){ //Browsercheck (needed)
        this.ver=navigator.appVersion;
        this.dom=document.getElementById?1:0;
        this.ie8=(this.ver.indexOf("MSIE 8")>-1 && this.dom)?1:0;
        this.ie7=(this.ver.indexOf("MSIE 7")>-1 && this.dom)?1:0;
        this.ie6=(this.ver.indexOf("MSIE 6")>-1 && this.dom)?1:0;
        this.ie5=(this.ver.indexOf("MSIE 5")>-1 && this.dom)?1:0;
        this.ie4=(document.all && !this.dom)?1:0;
        this.ie=this.ie4||this.ie5||this.ie6||this.ie7||this.ie8;
        this.ns5 = (navigator.vendor == ("Netscape6") || navigator.product == ("Gecko"));
        this.ns4=(document.layers && !this.dom)?1:0;
        this.ns=this.ns4||this.ns5;
        this.bw=(this.ie5 || this.ie4 || this.ns4 || this.ns5 || this.ie6);
        this.IECSS = (this.ie && document.compatMode) ? document.compatMode == "CSS1Compat" : false;
        this.IEDTD = (this.ie && document.doctype) ? document.doctype.name.indexOf(".dtd")!=-1 : this.IECSS;
        return this;
}

function findObj(n, d) {
  var p,i,x;
  x = 0;
  if(!d) {
    d=document;
  }
  p = n.indexOf("?");
  if((p > 0)&&(parent.frames.length)) {
    d=parent.frames[n.substring(p+1)];
    if (d) {
      d = d.document;
    }
    n=n.substring(0,p);
  }
  if (d.getElementById) {
    x = d.getElementById(n);
  }
  if (!x) {
    if (!(x=d[n])&&d.all) {
      x = d.all[n];
    }
  }
  if (parent.frames&!x) {
    x = parent.frames[n];
  }
  if (d.links&!x) {
    x=d.links[n];
  }
  if (!x&&d.forms) {
    for (i=0;!x&&i<d.forms.length;i++) {
      if (d.forms[i].name == n) {
        x = d.forms[i];
      }
      else {
        x=d.forms[i][n];
      }
    }
  }
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) {
    if (d.layers[i].name == n) {
      x = d.layers[i];
    }
    else if (d.layers[i].id == n) {
      x = d.layers[i];
    }
    else {
      x=findObj(n,d.layers[i].document);
    }
  }
  if (!x&&d.images) {
    if (!x) {
      x=d.images[n];
    }
    for (i=0; !x&&d.images&&i<d.images.length;i++) {
      if (d.images[i].name == n) {
        x = d.images[i];
      }
    }
  }
  return x;
}

function PopUp(Type, Country, Text, WindowName, Width) {
  var Doc, Contents;
  if (Width == null) {
	  Width = 540;
  }
  if (!Type) {
	  WindowName = Country;
  }
  if (WindowName == null) {
	  WindowName = "TextWindow";
  }
  if (Type == "EXTERNAL") {
  	testwin = open(Country, "External",'height=140,width=575,scrollbars=yes,toolbars=no,resizable=yes');
	testwin.focus();
  }
  else {
    Doc = findObj(WindowName, self.document);
    if (!Doc) {
      Doc = self.document.createElement('div');
      Doc.setAttribute('id', 'TextWindow');
      Doc.innerHTML = '&nbsp;';
      Doc.style.width= '540px';
      self.document.body.appendChild(Doc);
    }
    if (Doc&&Doc.innerHTML) {
      if (!Text) {
        Text = '';
      }
      if (Type) {
        if (Type != 'TEXT') {
	      Contents = '<TABLE CELLSPACING=0 CELLPADDING=1 BORDER=0 WIDTH='+Width+' BGCOLOR="#D3D3D3"><TR><TD COLSPAN=2><B>'+Country+'</B></TD><TD ALIGN="RIGHT"><A HREF="" onClick="PopUp(0, \''+WindowName+'\');return false;"><IMG HSPACE=2 WIDTH=16 HEIGHT=14 SRC="http://www.wipo.int/ipdl/include/close.gif" BORDER=0></A></TD></TR>'+eval(Type+"CommonMessage")+Text+'</TABLE>';
        }
        else {
	   	  if ((Country != "BODY") && (Country != 'NORMAL')) {
			  Contents = '<TABLE CELLSPACING=0 CELLPADDING=1 BORDER=0 WIDTH='+Width+' BGCOLOR="#D3D3D3"><TR><TD ALIGN="RIGHT"><A HREF="" onClick="PopUp(0, \''+WindowName+'\');return false;"><IMG HSPACE=2 WIDTH=16 HEIGHT=14 SRC="http://www.wipo.int/ipdl/include/close.gif" BORDER=0></A></TD></TR><TR><TD>'+Text+'</TD></TR></TABLE>';
		  }
		  else {
			  Contents = Text;
		  }
        }
		if ( !((Type == 'TEXT') && ((Country == 'BODY') || (Country == 'NORMAL'))) ){
			Doc.style.background = "#D3D3D3";
		}
        Doc.innerHTML = Contents;
		Doc.style.visibility = "visible";
	self.popped = true;
      }
      else {
	self.popped = false;
        Doc.innerHTML = '&nbsp;';
        Doc.style.background = "#FFFFFF";
		Doc.style.visibility = "hidden";
      }
    }
    else if (Doc) {
      Doc = Doc.document;
      OldDocHeight = Doc.height;
	  if (Type) {
        if (Type != 'TEXT') {
	      Doc.write('<TABLE CELLSPACING=0 CELLPADDING=3 BORDER=0 WIDTH='+Width+' BGCOLOR="#D3D3D3"><TR><TD COLSPAN=2 NOWRAP>'+Country+'</TD><TD ALIGN="RIGHT"><A HREF="" onClick="PopUp(0);return false;"><IMG SRC="http://www.wipo.int/ipdl/include/close.gif" WIDTH=16 HEIGHT=14 BORDER=0></A></TD></TR>');
	      eval("Doc.write("+Type+"CommonMessage)");
		self.popped = true;
        }
      }
      else {
	self.popped = false;
        Doc.write('&nbsp;');
        Doc.close();
        self.document.height -= OldDocHeight-(Doc.height);
        self.scrollBy(0, -1*(OldDocHeight));
      }
      if (Text) {
        if (Type != 'TEXT') {
	      Doc.write('<TR><TD COLSPAN=3>');
        }
        Doc.write(Text);
        if (Type != 'TEXT') {
	      Doc.write('</TD></TR>');
        }
	self.popped = true;
      }
      if (Type) {
		if (Type != 'TEXT') {
	        Doc.write('</TABLE>');
		}
        Doc.close();
        self.document.height -= (OldDocHeight + 5);
        self.document.height += (Doc.height + 5);
        self.scrollBy(0, Doc.height);
	self.popped = true;
      }
    }
  }
}
function setValue(Input, NewValue) {
  var p, Name, obj;
  p = Input.name.indexOf('_DUMMY');
  Name = Input.name.substring(0, p);
  obj = eval('self.document.frm.'+Name);
  if (Input.type == 'checkbox') {
    if (obj.value) {
      obj.value = '';
    }
    else {
      obj.value = NewValue;
    }
  }
  else {
    obj.value = NewValue;
  }
}

function copyValue(Input) {
  var obj, type;
  obj = findObj('DummyForm', self.document);
  obj = eval('obj.'+Input.name+'_DUMMY');
  if (!obj) {
    return;
  }
  if (!obj.type) {
    type = obj[0].type;
  }
  else {
    type = obj.type;
  }
  if (type == 'checkbox') {
    if (Input.value) {
      obj.checked = true;
    }
    else {
      obj.checked = false;
    }
  }
  else if (type.indexOf('select') > -1) {
    for (i = 0; i < obj.options.length; i++) {
      if (obj.options[i].value == Input.value) {
		obj.selectedIndex = i;
		break;
      }
    }
  }
  else if (type == 'radio') {
    for (i = 0; i < obj.length; i++) {
      if (obj[i].value == Input.value) {
	obj[i].checked = true;
	break;
      }
    }
  }
}

function setVisibility(obj, toggle) {
  obj = findObj(obj, self.document);
  if (!obj) {
    obj = eval('self.document.'+obj);
    return;
  }
  var Local = obj;
  if (Local.style) {
    Local = Local.style;
  }
  if (toggle == 1) {
    Local.visibility = 'visible';
  }
  else {
    Local.visibility = 'hidden';
  }
}
function fieldcodes() {
	testwin = open('about:blank','FieldCodes','height=400,width=640,scrollbars=yes,toolbars=no,resizable=yes');
	testwin.document.write(FieldCodeText);
	testwin.document.close();
	testwin.focus();
}
function samplesearch() {
	testwin = open('about:blank','Sample','height=475,width=485,scrollbars=yes,toolbars=no,resizable=yes');
	testwin.document.write(SampleText);
	testwin.document.close();
	testwin.focus();
}

function dateObj() {
  this.day = this.month = this.year = 0;
  this.validMonth = this.validDay = this.validYear = false;
  this.daysInMonth = new Array(0,31,28,31,30,31,30,31,31,30,31,30,31);
  this.languages = new Array("ENG", "FRE", "SPA");
  this.monthNames = new Array();
  this.monthNames[0] = new Array("January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December");
  this.monthNames[1] = new Array("janvier", "f&eacute;vrier", "mars", "avril", "mai", "juin", "j);uillet", "ao&ucirc;t", "septembre", "octobre", "novembre", "d&eacute;cembre");
  this.monthNames[2] = new Array("enero", "febrero", "marcha", "abril", "mayo", "junio", "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre");
y  
  this.dayNames = new Array();
  this.dayNames[0] = new Array("Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat");
  this.dayNames[1] = new Array("dim", "lun", "mar", "mer", "jue", "ven", "sam");
  this.dayNames[2] = new Array("dom", "lun", "mar", "mi&eacute;", "jue", "vie", "s&aacute;b");
  
  this.desc = "date";
  this.dateString = "";
  this.setLeapYear = function setLeapYear(input) {
	  if ((input %4 == 0) &&
		   ((input %100 != 0) || (input %400 == 0))
		   ) {
		  this.daysInMonth[2] = 29;
	  }
	  else {
		  this.daysInMonth[2] = 28;
	  }
  }
  this.setDateValues = function setDateValues(newDay, newMonth, newYear) {
	 this.day = newDay;
	 this.month = newMonth;
	 this.year = newYear;
	 this.date = 0;
	 this.setLeapYear(this.year);
	 if ((this.month >= 1) && (this.month <= 12)) {
  	  this.validMonth = true;
 	 }
 	 if ((this.day >= 1) && (this.day <= this.daysInMonth[this.month])) {
 	   this.validDay = true;
 	 }
 	 if (this.year != 0) {
  	   this.validYear = true;
 	 }
	 if ((this.validMonth) && (this.validDay) && (this.validYear)) {
		 this.date = (this.year*10000)+(this.month*100)+this.day;
	 }
  }
  this.addMonths = function addMonths(input) {
    var years = this.year;
    var months = parseInt(input)+this.month;
	if (months > 0) {
    	if (months > 12) {
			years += Math.floor((months-1)/12);
			months = ((months-1)%12)+1;
    	}
	}
	else {
		if (months < 1) {
			years += Math.ceil((months-12)/12);
			months = 12+(months%12);
		}
	}
	this.setLeapYear(years);
	if (this.day > this.daysInMonth[months]) {
		this.setDateValues(this.daysInMonth[months], months, years);
	}
	else {
		this.setDateValues(this.day, months, years);
	}
	this.desc += "+"+input+" months";
  }
  this.addDays = function addDays(input) {
	  var days = this.day + parseInt(input);
	  months = this.month;
	  years = this.year;
	  if (days > 0) {
		  // could just calculate this directly, technically.
		  while (days > this.daysInMonth[months]) {
			  days -= this.daysInMonth[months];
			  months++;
			  if (months > 12) {
				  years++;
				  this.setLeapYear(years);
				  months = 1;
			  }
		  }
	  }
	  else {
		  while (days < 1) {
			  months--;
			  if (months < 1) {
				  months = 12;
				  years--;
				  this.setLeapYear(years);
			  }
			  days += this.daysInMonth[months];
		  }
	  }
	this.setDateValues(days, months, years);
  }
  this.addYears = function addYears(input) {
	  var years = this.year+parseInt(input);
	  this.setLeapYear(years);
	  if (this.day > this.daysInMonth[this.month]) {
		  this.setDateValues(this.daysInMonth[this.month], this.month, years);
	  }
	  else {
		  this.setDateValues(this.day, this.month, years);
	  }
  }
  this.goToThursday = function goToThursday() {
	  var tempDate = new Date(this.year, this.month, this.day);
	  var currentDay = tempDate.getDay();
	  if (currentDay < 4) {
		this.addDays(4-currentDay);
	  }
	  else if (currentDay > 4) {
		this.addDays(6-(Math.floor(currentDay/6)));
	  }
  }
  this.goToWeekday = function goToWeekday() {
	  var tempDate = new Date(this.year, this.month, this.day);
	  var currentDay = tempDate.getDay();
	  if (currentDay == 0) {
		  addDays(1);
	  }
	  else if (currentDay == 6) {
		  addDays(2);
	  }
  }  
  this.setDateString = function setDateString(dateString, separator, format) {
	 var pos1 = dateString.indexOf(separator);
 	 var pos2= dateString.indexOf(separator, pos1+1);
 	 if (pos1==-1 || pos2==-1){
	   this.validMonth = this.validDay = this.validYear = false;
  	   return false;
 	 }
	 this.validMonth = this.validDay = this.validYear = true;
 	 var monthString, dayString, yearString;
 	 if (format.indexOf("MM") == 0) {
  	   monthString = dateString.substring(0, pos1);
  	   dayString = dateString.substring(pos1+1,pos2);
  	   yearString = dateString.substring(pos2+1);
 	 }
 	 else {
 	   dayString = dateString.substring(0, pos1);
 	   monthString = dateString.substring(pos1+1,pos2);
 	   yearString = dateString.substring(pos2+1);
 	 }
	 this.setDateValues(parseInt(dayString), parseInt(monthString), parseInt(yearString));
	 this.dateString = dateString;
	 return true;
  }
  this.isValidDate = function isValidDate() {
	if ((this.isValidDay() == false) || (this.isValidMonth() == false) || (this.isValidYear() == false)) {
		return false;
	}
	else {
		return true;
	}
  }
  this.isValidDay = function isValidDay() {
	  if ((this.day > 0) && (this.day <= this.daysInMonth[this.month])) {
		  return true;
	  }
	  else {
		  return false;
	  }
  }
  this.isValidMonth = function isValidMonth() {
	  if ((this.month > 0) && (this.month <= 12)) {
		  return true;
	  }
	  else {
		  return false;
	  }
  }
  this.isValidYear = function isValidYear() {
	  if (this.year > 0) {
		  return true;
	  }
	  else {
		  return false;
	  }
  }
  this.compare = function compare(input) {
	  if (this.date < input.date) {
		  return -1;
	  }
	  else if (this.date > input.date ) {
		  return 1;
	  }
	  else {
		  return 0;
	  }
  }
  this.displayDate = function displayDate(language) {
	  var tempDate = new Date(this.year, this.month, this.day);
	 ;
	  var monthLanguage = 0;
	  for (i = 0; i < this.languages.length; i++) {
		  if (this.languages[i] == language) {
			  monthLanguage = i;
			  break;
		  }
	  }
	  var output = this.dayNames[monthLanguage][tempDate.getDay()]+", "+this.day+" "
	  output += this.monthNames[monthLanguage][this.month-1];
	  output += ", "+this.year;
	  return output;
  }	  
  this.copyValues = function copyValues(input) {
	  this.setDateValues(input.day, input.month, input.year);
  }
  return this;
}

function replace(input, stringToSearchFor, stringToReplaceWith) {
  var output, pos, hit;
  output = input;
  pos = 0;
  while ((hit = output.indexOf(stringToSearchFor, pos)) > 0) {
    output = output.substring(0, hit) + stringToReplaceWith + Output.substring(hit+stringToSearchFor.length, output.length);
    pos = hit+stringToReplaceWith.length;
  }
  return output;
}


function displayNumber(input, separator) {
  if (!separator) {
    sep = ",";
  }
  else {
    sep = separator;
  }
  temp = input;
  for (i = input.length-3; i > 0; i -= 3) {
    temp = temp.substr(0, i)+sep+temp.substr(i);
  }
  return temp;
}

function selectRadio(name, value) {
  var obj = findObj(name);
  if (obj) {
    for (var i = 0; i < obj.length; i++) {
      if (obj[i].value == value) {
        obj[i].checked = true;
        break;
      }
    }
  }
}
function selectRadioById(name) {
  var obj = findObj(name);
  if (obj) {
    obj.checked = true;
  }
}
function toggleCheckBox(name) {
  var obj = findObj(name);
  if (obj.checked == true) {
    obj.checked = false;
  }
  else {
    obj.checked = true;
  }
}
function wipoPop(elem, msg, w, xOff, yOff) {
  var ns4 = document.layers;
  var ns6 = document.getElementById&&!document.all;
  var ie4 = document.all;
  var nfo;
  if (ns4) {
	nfo=document.Info;
  }
  else if (ns6) {
	nfo=document.getElementById("Info");
  }
  else if (ie4) {
	nfo=document.all.Info;
  }
  if (!nfo) {
    if (ie4 || ns6) { 
    	nfo = document.createElement('div');
    	nfo.setAttribute('id', 'Info');
    	nfo.innerHTML = '&nbsp;';
    	nfo.style.width= '200px';
    	document.body.appendChild(nfo);
    }
    else {
	return;
    }
  }
  if (ns6 || ie4) {
    nfo = nfo.style;
  }
  if (!elem) {
    nfo.visibility = "hidden";
    nfo.left = 0;
    return;
  }
  if ((!msg) || (msg.length == 0)) {
    return;
  }
  if (!w) {
    w = "200";
  }
  if (!xOff) {
    xOff = 0;
  }
  if (!yOff) {
    yOff = 15;
  }
  if (ns4) {
//	document.captureEvents(Event.MOUSEMOVE);
  }
  else {
	nfo.visibility="visible"
	nfo.display="none"
  }
  var content="<TABLE WIDTH='" + w + "' CELLPADDING=2 CELLSPACING=0 STYLE='border: 1px solid rgb(100,100,100)'" + "BGCOLOR='#eeeeee'><TD>" + msg + "</TD></TABLE>";
  var x = xOff;
  var y = yOff;
  while (elem.offsetParent != null) {
    y += elem.offsetTop;
    x += elem.offsetLeft;
    elem = elem.offsetParent;
  }
  nfo.top = y+"px";
  nfo.left = x+"px";
  self.status = x+","+y;
  if (ns4) {
	nfo.document.write(content);
	nfo.document.close();
	nfo.visibility = "visible"
  }
  if (ns6) {
	document.getElementById("Info").innerHTML = content;
	nfo.display = ''
  }
  if (ie4) {
	document.all("Info").innerHTML = content;
	nfo.display = ''
  }
}
