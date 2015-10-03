var selectedTab = 1;
var selectedLink = -1;
// should just make an object.
var distributionGraphs = new Array();
var distributionDataSets = new Array();
var distributionFirstLinks = new Array();
var distributionTimeOut;
var generalDistributionGraph;
var distributionVisible = false;
var highlightInfo;
var sectionCount;
var sectionDivisor = 4;
var sectionTotals;
var graphHeight = 232;
var graphBoxHeight = 335;
var graphWidth = 200;
var graphBoxWidth = 215;
var labelHeight = 25;
var maxSectionHitCount;
var hitTotal = 0;
var graphLocations = new Array();
var generalGraphLocation;
var barSeparatorColor = "#222";
var highlightColor = "#fff";
var currentGraph = -2;
var lowerLanguage = "en";
if (self.location.pathname.indexOf("/fr/") > -1) {
  lowerLanguage = "fr";
}
if (self.location.pathname.indexOf("/es/") > -1) {
  lowerLanguage = "es";
}
TabNames = new Array("NULL", "Status", "Desc", "Claims", "National", "History", "Docs");
imageLink = "foo";
drawingString = "NULL";
function blurLink(linkName) {
  obj = findObj(linkName);
  if (obj) {
    obj.blur();
  }
}
function highlightTab(index, on) {
  obj = findObj(TabNames[index]+"Tab");
  if (obj) {
    if (on) {
      obj.className += " tabHighlight";
    }
    else {
      obj.className = obj.className.replace(/ tabHighlight/, "");
    }
  }
}
function highlightLink(index) {
  if (selectedLink != -1) {
    obj = findObj("LinkImage"+selectedLink);
    if (obj) {
      obj.src = "/shared/images/pxl/transpxl.gif";
    }
  }
  obj = findObj("LinkImage"+index);
  if (obj) {
    obj.src = "/shared/images/icon/link.gif";
  }
  selectedLink = index;
}
function selectTab(index) {
  obj = findObj(TabNames[selectedTab]+"Tab");
  var contentHeight, contentWidth;
  if (obj) {
    obj.className = "tabOff";
  }
  obj = findObj(TabNames[selectedTab]+"Content");
  if (obj) {
    obj.style.visibility = "hidden";
  }
  if (parseInt(index) > 0) {
    selectedTab = index;
  }
  else {
    for (i = 1; i < TabNames.length; i++) {
      if (TabNames[i].toUpperCase() == index.toUpperCase()) {
	selectedTab = i;
	break;
      }
    }
  }
  type = TabNames[selectedTab];
  obj = findObj(type+"Tab", self.document);
  if (obj) {
    obj.className = "tabOn";
    obj.blur();
  }
  obj = findObj(type+"Content");
  if (obj) {
    obj.style.visibility = "visible";
    obj.style.zIndex = "100";
//    alert(obj.innerHTML);
  }
  obj = findObj(type+"Bottom");
  if (obj) {
    contentHeight = parseInt(obj.offsetTop);
  }
  else {
//    alert("Cannot find: "+type+"Bottom");
  }
  if (contentHeight != undefined) {
    obj = findObj("Contents");
    obj.style.height = contentHeight+50+"px";
  }
  obj = findObj("Contents");
  if (obj) {
    contentWidth = parseInt(obj.offsetWidth);
  }
  obj = findObj("tabs");
  if (obj) {
    obj.style.width = contentWidth+"px";
  }
  return false;
}

function hideDistributionInfo() {
  var obj = findObj("graphDiv");
  if (obj.style) {
    obj = obj.style;
  }
  obj.visibility = 'hidden';
  distributionVisible = false;
  if (currentGraph == -1) {
    showGeneralGraphIcon(1);
  }
  currentGraph = -2;
}
function showGeneralGraphIcon(state) {
  var obj = findObj("graphIconDiv");
  if (obj.style) {
    obj = obj.style;
  }
  if (state) {
    obj.visibility = 'visible';
  }
  else {
    obj.visibility = 'hidden';
  }
}

function showDistributionInfo(termNum) {
  distributionVisible = true;
  termFound = 0;
  index = 0;
  for (i = 0; i < distributionInfo.length; i++) {
    if (distributionInfo[i][0] == termNum) {
      termFound = 1;
      index = i;
      break;
    }
  }
  if (termFound == 0) {
    return;
  }
  if (!graphLocations[index]) {
    contentTop = 0;
    contentLeft = 0;
    obj = findObj("highlightTerm-"+termNum);
    while (obj.offsetParent != null) {
      contentTop += obj.offsetTop;
      contentLeft += obj.offsetLeft;
      obj = obj.offsetParent;
    }
    graphLocations[index] = new Array(contentTop, contentLeft);
  }
  graphObj = findObj("graphDiv");
  if (graphObj) {
    graphObj.style.top = graphLocations[index][0]-3+"px";
    graphObj.style.left = graphLocations[index][1]-3+"px";
    graphObj.style.position = "absolute";
  }

  if ((!distributionGraphs[index]) || (distributionGraphs[index][1].length == 0)) {
    generateDistributionGraph(index, termNum);
  }
  showDistributionGraph(index);
}

function showGeneralDistributionInfo() {
  if ((!generalDistributionGraph) || (generalDistributionGraph.length == 0)) {
    generateGeneralDistributionGraph();
  }
  if (currentGraph == -1) {
    return;
  }
  if (!generalGraphLocation) {
    contentTop = 0;
    contentLeft = 0;
    obj = findObj("generalGraphLink");
    while (obj.offsetParent != null) {
      contentTop += obj.offsetTop;
      contentLeft += obj.offsetLeft;
      obj = obj.offsetParent;
    }
    generalGraphLocation = new Array(contentTop, contentLeft);
  }
  graphObj = findObj("graphDiv");
  if (graphObj) {
    graphObj.innerHTML = "&nbsp;";
    graphObj.style.top = generalGraphLocation[0]-40+"px";
    graphObj.style.left = generalGraphLocation[1]+38+"px";
    graphObj.style.position = "fixed";
  }
  showDistributionGraph(-1);
}

function highlightTermLink(termNum, section, state) {
  obj = findObj("highlightTermLink-"+termNum);
  bar = findObj("bar-"+termNum+"-"+section);
  if (state == 1) {
    if (obj.style) {
      obj.style.borderColor = "#000";
      obj.style.borderStyle = "dashed";
      obj.style.borderWidth = 1+"px";
    }
    if (bar.style) {
      barSeparatorColor = bar.style.borderColor;
      bar.style.borderColor = "#ddd";
      bar.style.borderStyle = "solid";
      bar.style.borderWidth = 1+"px";
      bar.style.zIndex = 10000;
    }
  }
  else {
    if (obj.style) {
      obj.style.borderColor = "#fff";
      obj.style.borderStyle = "solid";
      obj.style.borderWidth = 1+"px";
    }
    if (bar.style) {
      bar.style.borderColor = barSeparatorColor;
      bar.style.borderStyle = "solid";
      bar.style.borderWidth = 1+"px";
      bar.style.zIndex = 1;
    }
  }
}

function highlightTermBars(termNum, state) {
  /*
    // unfortunately, this doesn't work with a dynamically generated
//   layer.  boo.
  var mysheet=document.styleSheets[document.styleSheets.length-1];
  var myrules=mysheet.cssRules? mysheet.cssRules: mysheet.rules;
  var targetrule = "notFound";
  for (i=0; i< myrules.length; i++){
    if (myrules[i].selectorText.toLowerCase()==".hl-"+termNum) {
      targetRule=myrules[i];
      break;
    }
  }
  if (state) {
    highlightColor = targetRule.style.borderColor;
    targetRule.style.borderColor = "#000";
    targetRule.style.borderStyle = "dashed";
  }
  else {
    targetRule.style.borderColor = highlightColor;
    targetRule.style.borderStyle = "solid";
  }
  */
  obj = findObj("highlightTermLink-"+termNum);
  if (state == 1) {
    if (obj.style) {
      obj = obj.style;
    }
    highlightColor = obj.borderColor;
    obj.borderColor = "#000";
    obj.borderStyle = "dashed";
    obj.borderWidth = 1+"px";
  }
  else {
    if (obj.style) {
      obj = obj.style;
    }
    obj.borderColor = highlightColor;
    obj.borderStyle = "solid";
  }
  if (currentGraph != -1) {
    return;
  }
  obj = findObj("graphDiv");
  objStyle = obj;
  if (obj.style) {
    objStyle = obj.style;
  }
  if (objStyle.visibility != "visible") {
    return;
  }
  var bars = obj.getElementsByTagName("DIV");
  for (i = 0; i < bars.length; i++) {
    if (bars[i].className == "graphBar hl-"+termNum) {
      if (state) {
	bars[i].style.borderColor = "#ddd";
	bars[i].style.zIndex = 10000;
      }
      else {
	bars[i].style.borderColor = barSeparatorColor;
	bars[i].style.zIndex = 1;
      }
    }
  }
}

function generateLinkStyles(count) {
  var colors = new Array('#ddf', '#ff4', '#90EE90', '#ddd', '#ecc');
  var i, output;

  output = "<STYLE TYPE='text/css' title='highlightStyle'>\n";
  for (i = 0; i < count; i++) {
    output += ".hl-"+(i+1)+"{\nbackground-color: "+colors[i%5]+";\nborder-color: "+colors[i%5]+";\n}\n";
  }
  output += "</STYLE>\n";
  return output;
}

function generateDistributionGraph(index, termNum) {
  //  alert('generating: '+termNum);
  /*
    // No way to do all the dynamic highlight & linking with plotkit
  var options = {
    "IECanvasHTC": "/plotkit/iecanvas.htc",
    "barOrientation": "horizontal",
    "colorScheme": PlotKit.Base.palette(PlotKit.Base.baseColors()[2]),
    "padding": {left: 0, right: 0, top: 0, bottom: 0},
    "drawYAxis": false,
    "xOriginIsZero": true,
    "yOriginIsZero": true
  };
  var layout = new PlotKit.Layout("line", options);
  */
  //  alert("TotalCount: "+distributionInfo[index][1]+", Highest Count: "+distributionInfo[index][2]+" = scale: "+scaleAmount);

  generateDataSet(index);

  /*
  layout.addDataset("Term: "+index, dataSet);
  //layout.addDataset("Term: "+index, [[0, 0], [1, 1], [2, 1.414], [3, 1.73], [4, 2]]);
  layout.evaluate();
  var canvas = MochiKit.DOM.getElement("graph");
  var plotter = new PlotKit.SweetCanvasRenderer(canvas, layout, options);
  plotter.render();
  */
  var top, width, termNum;
  scaleAmount = graphWidth/distributionInfo[index][2];
  height = Math.floor(graphHeight/(sectionCount/sectionDivisor));
  distributionGraphs[index] = new Array (termNum, "");
  distributionGraphs[index][1] = makeDistributionHeader("<A CLASS='hl hl-"+termNum+"' HREF='#a-"+termNum+"-1'>"+highlightInfo[index][2]+"</A>", highlightInfo[index][1], Math.floor(height));
  lastSection = 0;
  for (i = 0; i < distributionDataSets[index].length; i++) {
    width = Math.ceil(distributionDataSets[index][i][1]*scaleAmount);
    termNum = distributionInfo[index][0];
    top = Math.ceil(distributionDataSets[index][i][0]*height)+labelHeight;
    section = distributionDataSets[index][i][0];
    distributionGraphs[index][1] += makeBlankBars((section-(lastSection+1)), Math.floor(height));
    distributionGraphs[index][1] += startGraphLine(Math.floor(height));
    distributionGraphs[index][1] += makeDistributionBar(termNum, section, Math.floor(height), width, 3, distributionDataSets[index][i][1], distributionFirstLinks[index][section], "");
    distributionGraphs[index][1] += endGraphLine();
    lastSection = section;
  }
  distributionGraphs[index][1] += makeBlankBars(((sectionCount/sectionDivisor)-(lastSection+1)), Math.floor(height));
  distributionGraphs[index][1] += endGraph();
  distributionGraphs[index][1] += makeDistributionFooter(distributionInfo[index][2]);
}

function generateDataSet(index) {
  if (distributionDataSets[index]) {
    return;
  }
  distributionDataSets[index] = new Array();
  var mapPos, currentInt, currentPos, mask, intPos, section, i, j, currentTotal;
  countPos = 0;
  section = 0;
  if ((sectionDivisor != 2) && (sectionDivisor != 4) && (sectionDivisor != 8) && (sectionDivisor != 16)) {
    sectionDivisor = 1;
  }
  if (!sectionTotals) {
    sectionTotals = new Array(Math.floor(sectionCount/sectionDivisor));
    for (i = 0; i < Math.ceil(sectionCount/sectionDivisor); i++) {
      sectionTotals[i] = 0;
    }
  }
  distributionFirstLinks[index] = new Array(Math.ceil(sectionCount/sectionDivisor)+1);
  currentTotal = 1;
  for (mapPos = distributionInfo[index][3].length-8; mapPos >= 0; mapPos -= 8) {
    currentInt = parseInt(distributionInfo[index][3].substring(mapPos, mapPos+8), 16);
    mask = 1;
    for (intPos = 0; intPos < 32; ) {
      currentCount = 0;
      for (j = 0; j < sectionDivisor; j++) {
	if ((currentInt & mask) != 0) {
	  currentCount += parseInt(distributionInfo[index][4].substring(countPos, countPos+2), 16);
	  countPos += 2;
	}
	currentInt >>= 1;
	intPos++;
      }
      distributionFirstLinks[index][section] = currentTotal;
      if (currentCount > 0) {
	distributionDataSets[index][distributionDataSets[index].length] = new Array(Math.floor(section), currentCount);
	sectionTotals[section] += currentCount;
	hitTotal += currentCount;
	currentTotal += currentCount;
      }
      section++;
    }
  }
  if (sectionDivisor != 1) {
    for (i = 0; i < distributionDataSets[index].length; i++) {
      if (distributionDataSets[index][i][1] > distributionInfo[index][2]) {
	distributionInfo[index][2] = distributionDataSets[index][i][1];
      }
    }
  }
}

function generateGeneralDistributionGraph() {
  // We need dataSets for all existing distributions
  if (generalDistributionGraph) {
    //    alert('already done');
    return;
  }
  for (i = 0; i < distributionInfo.length; i++) {
    generateDataSet(i);
  }
  distPos = new Array(distributionDataSets.length);
  for (i = 0; i < distributionDataSets.length; i++) {
    distPos[i] = 0;
  }
  // maxSectionHitCount should be set already.
  // with the possibility of combining sections dynamically, we again
  // have to check.
  var max = 0;
  for (i = 0; i < sectionTotals.length; i++) {
    if (sectionTotals[i] > max) {
      max = sectionTotals[i];
    }
  }
  scaleAmount = graphWidth/max;
  height = Math.floor((graphHeight)/(sectionCount/sectionDivisor));
  var sectionFound, offset, width, minSection, top;
  generalDistributionGraph = makeDistributionHeader("<A HREF='#' STYLE='text-decoration: none;'>General Term Distribution</A>", hitTotal, Math.floor(height));
  lastSection = 0;
  while (true) {
    sectionFound = false;
    minSection = (sectionCount/sectionDivisor)+1;
    for (i = 0; i < distributionDataSets.length; i++) {
      if ((distPos[i] < distributionDataSets[i].length) && (distributionDataSets[i][distPos[i]][0] < minSection)) {
	minSection = distributionDataSets[i][distPos[i]][0];
	sectionFound = true;
      }
    }
    if (sectionFound == false) {
      break;
    }
    offset = 3;
    generalDistributionGraph += makeBlankBars((minSection-(lastSection+1)), Math.floor(height));
    generalDistributionGraph += startGraphLine(Math.floor(height));
    lineDataSet = new Array();
    lineCount = 0;
    for (i = 0; i < distributionDataSets.length; i++) {
      if ((distPos[i] < distributionDataSets[i].length) && (distributionDataSets[i][distPos[i]][0] == minSection)) {
	// we show this one.
	width = Math.ceil(distributionDataSets[i][distPos[i]][1]*scaleAmount);
	termNum = distributionInfo[i][0];
	generalDistributionGraph += makeDistributionBar(termNum, minSection, height, width, offset, distributionDataSets[i][distPos[i]][1], distributionFirstLinks[i][minSection], lineCount, highlightInfo[i][2]);
	offset += width;
	distPos[i]++;
	lineCount++;
      }
    }
    generalDistributionGraph += endGraphLine();
    lastSection = minSection;
  }
  generalDistributionGraph += makeBlankBars((sectionCount/sectionDivisor)-(lastSection+1), Math.floor(height));
  generalDistributionGraph += endGraph();
  generalDistributionGraph += makeDistributionFooter(max);
}  

function endGraphLine() {
  var output = "</DIV>";
  return output;
}

function endGraph() {
  return(endGraphLine());
}

function makeDistributionBar(termNum, section, height, width, offset, hits, firstLink, barCount, term) {
  var output = "<DIV ID='bar-"+termNum+"-"+section+"' CLASS='graphBar hl-"+termNum+"' STYLE='position: absolute; left: "+offset+"px; width: "+width+"px; height: "+height+"px; font-size: "+height+"px;'><A TITLE='";
  /*
  var output = "<DIV ID='bar-"+termNum+"-"+section+"' CLASS='graphBar hl-"+termNum+"' STYLE='position: relative; ";
  if (barCount > 0) {
    output += "margin-left: -1px; ";
  }
  output += "display: inline; width: "+width+"px; height: "+height+"px; font-size: "+height+"px;'><A TITLE='";
  */
  if (term && (term.length > 0)) {
    output += term+" ("+hits+")";
  }
  else {
    output += hits;
  }
  output += "' CLASS='graphBar' HREF='#a-"+termNum+"-"+firstLink+"' onMouseOver='highlightTermLink("+termNum+", "+section+", 1);' onMouseOut='highlightTermLink("+termNum+", "+section+", 0);'><IMG BORDER=0 HEIGHT="+height+" width="+width+" src='/shared/images/pxl/transpxl.gif' HSPACE=0 VSPACE=0></A></DIV>";
  return output;
}

function makeDistributionHeader(Text, hits, height) {
  var output;
  output = "<DIV STYLE='background-color: #eee; border: #000 solid 1px; padding: 2px; margin-top: 0px; margin-bottom: 0px;'>"+Text+" ("+hits+")</DIV>";
  output += "<DIV STYLE='z-index: 1200; position: absolute; top: 3px; left: 192px;'><A HREF='javascript:hideDistributionInfo();'><IMG STYLE='float: right;' SRC='http://www.wipo.int/ipdl/include/close.gif' BORDER=0 HEIGHT='14' WIDTH='16' HSPACE=3></A></DIV>";
  output += "<DIV STYLE='border: #000 solid 1px; border-top: none; margin-left: 0px; padding: 3px; font-size: "+height+"px;'>";
  return output;
}

function makeDistributionFooter(max) {
  var output = "<DIV STYLE='height: "+(labelHeight-10)+"px; background-color: #eee; border: #000 solid 1px; border-top: none; margin-top: 0px !important; margin-top: -1px;'><DIV STYLE='float: left; padding-left: 3px;'>0</DIV><DIV STYLE='text-align: right;'>"+max+"&nbsp;</DIV></DIV>";
  return output;
}

function makeBlankBars(count, height) {
  var output = "";
  var i;
  for (i = 0; i < count; i++) {
    output += "<DIV CLASS='emptyBar' STYLE='height: "+height+"px; font-size: "+(height-3)+"px;'><DIV STYLE='margin: 0px;'><IMG HEIGHT="+height+" width="+height+" SRC='/shared/images/pxl/transpxl.gif' VSPACE=0 HSPACE=0></DIV></DIV>";
  }
  return output;
}

function startGraphLine(height) {
  var output = "<DIV CLASS='barContainer' STYLE='height: "+height+"px; font-size: "+height+"px;'>";
  return output;
}
function setInnerHTML(obj, text) {
  obj.innerHTML = text;
  obj.style.visibility = "visible";
}
function showDistributionGraph(index) {
  if (currentGraph == index) {
    return;
  }
  if (index > -2) {
    Doc = findObj("graphDiv");
    if (Doc.innerHTML) {
      Doc.style.visibility = "hidden";
      if (index > -1) {
	setInnerHTML(Doc, distributionGraphs[index][1]);
      }
      else {
	setTimeout("setInnerHTML(Doc, generalDistributionGraph)", 10);
      }
    }
    else {
      Doc.visibility = "hidden";
      if (index > -1) {
	Doc.document.write(distributionGraphs[index][1]);
      }
      else {
	Doc.document.write(generalDistributionGraph);
      }
      Doc.visibility = "visible";
    }
    currentGraph = index;
    if (currentGraph == -1) {
      showGeneralGraphIcon(0);
    }
    else {
      showGeneralGraphIcon(1);
    }
  }
}
function showHighlightInfo(input) {
  var Doc;
  var output = '';
  if (input.length > 0) {
    for (i = 0; i < input.length; i++) {
      if (input[i][1] > 0) {
	output += "<A NAME='a-"+input[i][0]+"-0'></A>";
      }
    }
    output += generateLinkStyles(input[input.length-1][0]);
    Doc = findObj("tabTop");
    if (Doc) {
      if (Doc.innerHTML) {
	Doc.innerHTML = output;
      }
      else {
	Doc.document.write(output);
      }
    }
  }
  var termNum, termCount, term;
  if (input.length > 0) {
    output ='<DIV ID="TermList">';
    if (lowerLanguage == "fr") {
      output += "Les termes de recherche suivants sont mis en &eacute;vidence dans ce document: ";
    }
    else if (lowerLanguage == "en") {
      output += "The following query terms are highlighted in this document: ";
    }
    else if (lowerLanguage == "es") {
      output += "Los t&eacute;rminos de la b&uacute;squeda son puestos en evidencia en le presente documento: ";
    }
    for (i = 0; i < input.length; i++) {	
      termCount = input[i][1];
      if (termCount) {
	termNum = input[i][0];
	term = input[i][2];
	output += "<SPAN ID='highlightTerm-"+termNum+"'><A ID='highlightTermLink-"+termNum+"' CLASS='hl hl-"+termNum+" hlTop' HREF='#' TITLE='"+termCount+"' onMouseOver='highlightTermBars("+termNum+", 1);' onMouseOut = 'highlightTermBars("+termNum+", 0);' onClick='showDistributionInfo("+termNum+"); return false;'>"+term+"</A>&nbsp;</SPAN>";
      }
    }
    output += "</DIV>";
    output += "<DIV ID='graphIconDiv'><A onMouseOver='wipoPop(this, \"";
    output += "Click the <b>Distribution of terms</b> icon to see the quantity and location of searched for terms in relation to each other in each Description or Claims tab. The total number of all terms found is stated at the top of the chart.  Hovering over each specified colored block will:<OL><LI>highlight and state which term the block relates to and</LI><LI>state the number of nearby occurrences of the specific term.</OL>Clicking a colored block will take you to a nearby relatively located occurrence of the term in the full text.  That is, clicking a term in the middle of the chart will take you to a term in the middle of the full text.  Clicking that term will take you to the next occurrence of the term and so on...";
    output += "\", 300, -200, 25);' onMouseOut='wipoPop();' CLASS='graphLink' ID='generalGraphLink' HREF='javascript:showGeneralDistributionInfo();'><IMG SRC='/shared/images/icon/graph-icon.gif' height=30 width=24 border=0></A></DIV>";
    Doc = findObj("Highlights");
    if (Doc) {
      if (Doc.innerHTML) {
	Doc.innerHTML = output;
      }
      else {
	Doc.document.write(output);
      }
    }
  }
}
function customInit() {
  sortables_init();
}
function bookmark(url,title) {
  if ((navigator.appName == "Microsoft Internet Explorer") && (parseInt(navigator.appVersion) >= 4)) {
  window.external.AddFavorite(url,title);
  } else if (navigator.appName == "Netscape") {
    window.sidebar.addPanel(title,url,"");
  } else {
    alert("Press CTRL-D (Netscape) or CTRL-T (Opera) to bookmark");
  }
}
function linkToJP(docNum, docNumEn, sysTime, sysTimeEn) {
  $('jpDocNum').value = docNum;
  $('jpDocNumEn').value = docNumEn;
  $('jpSysTime').value = sysTime;
  $('jpSysTimeEn').value = sysTimeEn;
  $('jpLinkForm').submit();
}
