inSelectBox = false;
sfHover = function() {
  var menus = document.getElementsByTagName("UL");
  for (i = 0; i < menus.length; i++) {
    if (menus[i].className.indexOf('dropDown') > -1) {
      var items = menus[i].getElementsByTagName("LI");
      for (var j=0; j < items.length; j++) {
	items[j].onmouseover=function() {
	  if (this.className.indexOf("sfhover") == -1) {
	    this.className+=" sfhover";
	  }
	}
	items[j].onmouseout=function() {
	  if (inSelectBox == false) {
	    this.className=this.className.replace(new RegExp(" sfhover\\b"), "");
	  }
	}
	var subMenus = items[j].getElementsByTagName("UL");
	if (subMenus && subMenus.length > 0) {
	  var subItems = subMenus[0].getElementsByTagName("LI");
	  if (subItems && subItems.length > 0) {
	    subItems[0].className += " first";
	  }
	}
      }
      selects = menus[i].getElementsByTagName("SELECT");
      for (var k = 0; k < selects.length; k++) {
	selects[k].onclick=function() {
	  inSelectBox = true;
	}
	selects[k].onblur=function() {
	  inSelectBox = false;
	}
	selects[k].onchange=function() {
	  setTimeout("inSelectBox = false;", 25);
	}
      }
    }
  }
}  

if (window.attachEvent) {
  window.attachEvent("onload", sfHover);
}

highlightNavLink = function() {
  var myAddress;
  if ((typeof navParent != "undefined") && (navParent != null)) {
    myAddress = navParent;
  }
  else {
    myAddress = self.location.href;
  }
  var menu = $('l-col');
  if (menu != null) {
    var menuLinks = menu.getElementsBySelector('li');
    for (var i=0; i < menuLinks.length; i++) {
      var thisLink = menuLinks[i].getElementsByTagName("a");
      if (thisLink.length > 0) {
	if (thisLink[0] == myAddress) {
	  menuLinks[i].addClassName('navHighlight');
	}
      }
    }
  }
}
