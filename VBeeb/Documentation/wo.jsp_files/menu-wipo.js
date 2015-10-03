languageArray = new Array("en","es","fr","ru","ko","zh","ar","de","ja","pl");
var menuText = '';
var menuInitialized = 0;
var exclude = false;
var mu = false;
var dropMenuItemOpen = false;
var subMenuOpen = false;
if (typeof currentSection == "undefined") {
	currentSection = "resource";
}
function SP() {
}
function startMenu() {
	menuText = "<ul id='udm' class='udm'>\n<li class='udmHeader'>";
}
function startSubMenu() {
	menuText += "<ul>";
	subMenuOpen = true;
}
function endDropMenuItem() {
	if (subMenuOpen == true) {
		menuText += "</li>\n";
	}
	else if (dropMenuItemOpen == true) {
		menuText += "</li>\n";
		dropMenuItemOpen = false;
	}
}
function endSubMenu() {
	if (subMenuOpen == true) {
		menuText += "</ul>";
	}
	subMenuOpen = false;
}
function MI(url, linkText) {
	if (menuText.length == 0) {
		startMenu();
	}
	else {	
		endDropMenuItem();
		menuText += "</ul>\n</li>\n<li class='udmHeader'>";
	}
	menuText += "<a class='udmHeader";
	if (url.indexOf("http://") > -1) {
		currentCategory = url.substring(url.indexOf("/", 8)+1);
	}
	else {
		currentCategory = url.substring(1);
	}
	currentCategory = currentCategory.substring(0, currentCategory.indexOf("/"));
	if (currentCategory == currentSection) {
		menuText += " currentSection";
	}
	menuText+= "' href='http://www.wipo.int"+url+"'><span class='udmHeader'>"+linkText+"</span></a>\n<ul>";
}
function SI(url, linkText, hasSubMenu) {
	endDropMenuItem();
	menuText += "<li><a ";
	if (hasSubMenu == true) {
		menuText += "class='subMenuLink' ";
	}
	menuText += "href='http://www.wipo.int"+url+"'>"+linkText+"</a>";
	dropMenuItemOpen = true;
}
function endMenu() {
	endDropMenuItem();
	menuText += "</ul>\n</li>\n</ul>";
	menuInitialized = 1;
}
function writeMenu() {
	if (menuInitialized == 0) {
		endMenu();
	}
	try {
		document.write(menuText);
	}
	catch (failed) {
		localOnLoad = null;
		if (typeof self.onLoad == "function") {
			localOnLoad = window.onLoad;
		}
		self.onLoad = function() {
			Doc = document.createElement('div');
      			Doc.innerHTML = 'menuText';
      			document.body.appendChild(Doc);
			if (localOnLoad != null) {
				localOnLoad();
			}
		}
	}
}
function switchLanguageUrl(newLang, thisLang) {
	var localLanguage = "en";
	var url = self.document.location.href;
	if (thisLang) {
	  localLanguage = thisLang;
	}
	else {
		for (i = 0; i < languageArray.length; i++) {
		   if (url.indexOf("/"+languageArray[i]+"/") > 0) {
			localLanguage = languageArray[i];
			break;
		   }
		}
	}
	url = url.replace('/'+localLanguage+'/', '/'+newLang+'/');
	self.document.location.href = url;
}