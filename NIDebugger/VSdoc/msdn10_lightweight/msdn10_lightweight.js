/***********************************************
 * PAGE INIT
 ***********************************************/ 


// Event handler attachment
function registerEventHandler (element, event, handler) {
	if (element.attachEvent) {
		element.attachEvent('on' + event, handler);
	} else if (element.addEventListener) {
		element.addEventListener(event, handler, false);
	} else {
		element[event] = handler;
	}
}


// Event handler detachment
function unregisterEventHandler(element, event, handler) {
  if (typeof element.removeEventListener == "function")
    element.removeEventListener(event, handler, false);
  else
    element.detachEvent("on" + event, handler);
}


//element manipulation
//window.onload=init;
registerEventHandler(window, 'load', init);
		


function init()
{
	try {
		fixMoniker();
		fixMsdnLinks();
		mergeCodeSnippets();
		loadLangFilter();
		showSelectedLanguages();
	}
	catch (e) {
		//alert(e + "/n" + e.stack);
	}
}

/***********************************************
 * END PAGE INIT
 ***********************************************/ 


/**********************************************
 * EXPAND / COLLAPSE SECTION
 **********************************************/ 

function toggleSection(sectionLinkElm) {
	var sectionDiv = sectionLinkElm.parentNode.parentNode.parentNode;
	if (hasElementClass(sectionDiv, "collapsed")) {
		expandSection(sectionLinkElm);
	} else {
		collapseSection(sectionLinkElm);
	}
}

function expandSection(sectionLinkElm) {
	var sectionDiv = sectionLinkElm.parentNode.parentNode.parentNode;
	removeClassFromElement(sectionDiv, "collapsed");
	sectionLinkElm.setAttribute("title", "Collapse");	
}

function collapseSection(sectionLinkElm) {
	var sectionDiv = sectionLinkElm.parentNode.parentNode.parentNode;
	addClassToElement(sectionDiv, "collapsed");
	sectionLinkElm.setAttribute("title", "Expand");	
}

/**********************************************
 * END EXPAND / COLLAPSE SECTION
 **********************************************/ 



/***********************************************
 * CODE SNIPPETS
 ***********************************************/ 

/**
 * Merges adjacent code snippets in different languages into
 * single code collection with tabs.  
 */ 
function mergeCodeSnippets() {
	var allNodes = getElementAndTextNodes(document.body);
	var parentCodeSnippet = null;
	var i;
	
	for (i=0; i<allNodes.length; i++) {
		var currentNode = allNodes[i];
		
		if (!parentCodeSnippet) {
			// look for the first code snippet which will be a parent
			if (hasElementClass(currentNode, "codeSnippetContainer")) {
				parentCodeSnippet = currentNode;
				// snippet found, move after it
				var nextNode = getNextNonChildElementOrTextNode(parentCodeSnippet);
				while (allNodes[++i]!== nextNode && i<allNodes.length) {}
				i--;
			}
		
		} else {
			// look for the next ADJACENT code snippet
			if (hasElementClass(currentNode, "codeSnippetContainer")) {
				// merge it with the parent
				mergeTwoCodeSnippets(parentCodeSnippet, currentNode);
				// move after adjacent snippet
				var nextNode = getNextNonChildElementOrTextNode(currentNode);
				while (allNodes[++i]!== nextNode && i<allNodes.length) {}
				i--;				
			} else {
				// or look for the non-whitespace text
				if (currentNode.nodeType==TEXT_NODE || currentNode.nodeType==CDATA_SECTION_NODE) {
					if (currentNode.nodeValue.trim()!="") {
						// found non empty text after parent snippet, don't merge and find next parent
						parentCodeSnippet = null;
					}
				}
			}
		}
		
	}
}

var ELEMENT_NODE = 1;
var TEXT_NODE = 3;
var CDATA_SECTION_NODE = 4;

/**
 * Returns all elements and text nodes under the root.
 * Recursion is not used due to performance reasons. 
 */ 
function getElementAndTextNodes(root) {
    var result = [];

    var node = root.childNodes[0];
    while(node != null) {
				switch (node.nodeType) {
					case ELEMENT_NODE:
					case TEXT_NODE:
					case CDATA_SECTION_NODE:
						result.push(node);
						break;
				}    

        if(node.hasChildNodes()) {
            node = node.firstChild;
        }
        else {
            while(node.nextSibling == null && node != root) {
                node = node.parentNode;
            }
            node = node.nextSibling;
        }
    }
    
    return result;
}


/**
 * Merges two separate snippets together. The child snippet
 * becomes a part of the parent snippet. Tabs and visibility 
 * are adjusted accordingly.
 */  
function mergeTwoCodeSnippets(parentSnippet, childSnippet) {
	var childCode = getDivWithClass(childSnippet, "codeSnippetCode");
	var parentCodeCollection = getDivWithClass(parentSnippet, "codeSnippetCodeCollection");
	if (childCode && parentCodeCollection) {
		// remove existing lang code in parent, if any
		var lang = getLangOfCodeSnippetCode(childCode);
		if (lang) {
			var existingCode = getCodeSnippetCodeByLang(parentSnippet, lang);
			if (existingCode) {
				existingCode.parentNode.removeChild(existingCode);
			}
		}
		// move child code to the parent
		childCode.parentNode.removeChild(childCode);
		parentCodeCollection.appendChild(childCode);
		showHideTag(childSnippet, false);
		
		// correct the tabs (bold or normal for N/A lang)
		var csCode, tab, i;
		var tabsDiv = getDivWithClass(parentSnippet, "codeSnippetTabs");
		var langs = ["codeVB", "codeCsharp", "codeCpp", "codeFsharp", "codeJScript"];
		
		for (i=0; i<langs.length; i++) {
			lang = langs[i];
			csCode = getCodeSnippetCodeByLang(parentCodeCollection, lang);
			tab = getDivWithClass(tabsDiv, lang);
			if (csCode) {
				// lang exists
				removeClassFromElement(tab, "csNaTab");
			} else {
				// lang doesn't exist
				addClassToElement(tab, "csNaTab");
			}
		}
	}
}


/**
 * Gets a language of DIV with codeSnippetCode class.
 * @return  "codeVB", "codeCsharp", "codeCpp", "codeFsharp", "codeJScript" or null.
 */ 
function getLangOfCodeSnippetCode(elm) {
	if (hasElementClass(elm, "codeVB")) {
		return "codeVB";
	} else if (hasElementClass(elm, "codeCsharp")) {
		return "codeCsharp";
	} else if (hasElementClass(elm, "codeCpp")) {
		return "codeCpp";
	} else if (hasElementClass(elm, "codeFsharp")) {
		return "codeFsharp";
	} else if (hasElementClass(elm, "codeJScript")) {
		return "codeJScript";
	} else {
		return null;
	}
}


/**
 * Gets a DIV with codeSnippetCode class with specified language class.
 * @return null if not found.
 */  
function getCodeSnippetCodeByLang(containerSnippet, lang) {
	var divTags = containerSnippet.getElementsByTagName("div");
	var i;
	for (i=0; i<divTags.length; i++) {
		if (hasElementClasses(divTags[i], new Array("codeSnippetCode", lang))) {
			return divTags[i];
		}
	}
	return null;
}


/**
 * Gets the next node which is not a child of specified element and
 * is whether an element or text node.
 * @remark Unlike nextSibling property, this method returns also text nodes.
 * Moreover, if there is no next sibling, this metod goes higher in the hierarchy
 * and finds the next node.
 */     
function getNextNonChildElementOrTextNode(nod) {
	// try next sibling first
	var res = nod;
	while ((res = res.nextSibling) != null) {
		switch (res.nodeType) {
			case ELEMENT_NODE:
			case TEXT_NODE:
			case CDATA_SECTION_NODE:
				return res;
				break;
		}
	}
	
	// no sibling element or text found, try higher
	if (nod.parentNode) {
		return getNextNonChildElementOrTextNode(nod.parentNode);
	} else {
		// no node found
		return null;
	}
}


/**
 * Returns the first DIV element which has the specified class.
 * @return null if not found. 
 */ 
function getDivWithClass(parentElm, className) {
	var divTags = parentElm.getElementsByTagName("div");
	var i;
	for (i=0; i<divTags.length; i++) {
		if (hasElementClass(divTags[i], className)) {
			return divTags[i];
		}
	}
	return null;
}


/**
 * Determines whether an element contains specified CSS class.
 */ 
function hasElementClass(elm, className) {
	if (elm.className) {
		var classes = elm.className.split(" ");
		className = className.toLowerCase();
		var i;
		for (i=0; i<classes.length; i++) {
			if (classes[i].toLowerCase() == className) {
				return true;
			}
		}
	}
	
	return false;
}


/**
 * Determines whether an element contains all specified CSS classes.
 * @param  classNames An array of class names.
 */ 
function hasElementClasses(elm, classNames) {
	if (elm.className) {
		var classes = elm.className.split(" ");		
		var i, j, found;
		found=0;
		for (j=0; j<classNames.length; j++) {
			var className = classNames[j].toLowerCase(); 
			for (i=0; i<classes.length; i++) {
				var elmClass = classes[i].toLowerCase();
				if (elmClass == className) {
					found++;
					break;
				}
			}
		}
		if (found==classNames.length) {
			return true;
		}
	}
	
	return false;
}

 
/**
 * Removes specified CSS class from an element, if any.
 */ 
function removeClassFromElement(elm, className) {
	if (elm == null) return;
	if (elm.className) {
		var classes = elm.className.split(" ");
		className = className.toLowerCase();
		var i;
		for (i=classes.length-1; i>=0; i--) {
			if (classes[i].toLowerCase() == className) {
				classes.splice(i, 1);
			}
		}
		elm.className = classes.join(" ");
	}
}


/**
 * Adds specified CSS class to an element.
 */ 
function addClassToElement(elm, className) {
	if (elm == null) return;
	if (elm.className) {
		var classes = elm.className.split(" ");
		var classNameLow = className.toLowerCase();
		var i;
		for (i=classes.length-1; i>=0; i--) {
			if (classes[i].toLowerCase() == classNameLow) {
				// class already exists
				return;
			}
		}
		classes[classes.length] = className;
		elm.className = classes.join(" ");
	} else {
		elm.className = className;
	}
}

/**
 * Superfast trim. Faster than pure regex solution.
 */ 
String.prototype.trim = function () {
	var	str = this.replace(/^\s\s*/, ''),
		ws = /\s/,
		i = str.length;
	while (ws.test(str.charAt(--i)));
	return str.slice(0, i + 1);
}

/***********************************************
 * END CODE SNIPPETS
 ***********************************************/ 



/***********************************************
 * LANGUAGE FILTER
 ***********************************************/ 
    
var languageToShow = "codeVB";    

function loadLangFilter() {
	languageToShow = loadSetting("languageToShow", "codeVB");
}


function saveLangFilter() {
	saveSetting("languageToShow", languageToShow);
}


/**
 * Hides/shows the language sections according to language filter
 * @param langCode  "VB", "Csharp", "Cpp", "Fsharp", "JScript"
 */ 
function CodeSnippet_SetLanguage(langCode) {
	languageToShow = "code" + langCode;
	showSelectedLanguages();
	saveLangFilter();
}


/**
 * A snippet is a DIV with class="codeSnippetContainer".
 */ 
function getAllCodeSnippets() {
	var divTags = document.getElementsByTagName("div");
	var snippets = new Array();
	var i, j;
	j = 0;
	for (i=0; i<divTags.length; i++) {
		if (hasElementClass(divTags[i], "codeSnippetContainer")) {
			snippets[j++] = divTags[i];
		}
	}
	return snippets;
}


/**
 * Hides/shows the language sections according to language filter
 */ 
function showSelectedLanguages() {
	try {
		var snippets = getAllCodeSnippets();
		var i, j, divs, codeCollection, snippet;
		
		for (i=0; i<snippets.length; i++) {
			snippet = snippets[i];
			if (snippet.style.display != "none") {
				var langIsNA = false;
			
				// set the tabs (active/inactive)
				var tabsDiv = getDivWithClass(snippet, "codeSnippetTabs");
				// reset corners
				var leftCorner, rightCorner;
				leftCorner = getDivWithClass(tabsDiv, "codeSnippetTabLeftCorner");
				if (!leftCorner) {
					leftCorner = getDivWithClass(tabsDiv, "codeSnippetTabLeftCornerActive");
				}
				rightCorner = getDivWithClass(tabsDiv, "codeSnippetTabRightCorner");
				if (!rightCorner) {
					rightCorner = getDivWithClass(tabsDiv, "codeSnippetTabRightCornerActive");
				}
				removeClassFromElement(leftCorner, "codeSnippetTabLeftCornerActive");
				addClassToElement(leftCorner, "codeSnippetTabLeftCorner");
				removeClassFromElement(rightCorner, "codeSnippetTabRightCornerActive");
				addClassToElement(rightCorner, "codeSnippetTabRightCorner");

				//  get the tabs
				divs = tabsDiv.getElementsByTagName("div");
				var tab;
				var tabDivs = new Array();
				for (j=0; j<divs.length; j++) {
					if (hasElementClass(divs[j], "codeSnippetTab")) {
						// it's a tab
						tab = divs[j];
						tabDivs[tabDivs.length] = tab;
					}
				}
				
				//  activate/deactivate the tabs
				var visibleTabs = new Array();
				for (j=0; j<tabDivs.length; j++) {
					tab = tabDivs[j];
					var tabLink;
					
					if (hasElementClass(tab, languageToShow)) {
						addClassToElement(tab, "csActiveTab");
						tabLink = tab.getElementsByTagName("a")[0];
						tabLink.removeAttribute("href");
						langIsNA = hasElementClass(tab, "csNaTab");
					} else {
						removeClassFromElement(tab, "csActiveTab");
						var shortLang = getLangOfCodeSnippetCode(tab).substring(4);
						tabLink = tab.getElementsByTagName("a")[0];
						tabLink.setAttribute("href", "javascript: CodeSnippet_SetLanguage('" + shortLang + "');");
					}
				
					//  get visible tabs; invisible tabs are: with not supported lang AND not active
					if (!(hasElementClass(tab, "csNaTab") && !hasElementClass(tab, "csActiveTab"))) {
						// tab is visible
						visibleTabs[visibleTabs.length] = tab;
					}
				}

				// fix some styles (first, last) of visible tabs and corners
				for (j=0; j<visibleTabs.length; j++) {
					tab = visibleTabs[j];
					var tabLink;
					
					removeClassFromElement(tab, "csFirstTab");
					removeClassFromElement(tab, "csLastTab");
					if (j == 0) {
						addClassToElement(tab, "csFirstTab");
						if (hasElementClass(tab, "csActiveTab")) {
							removeClassFromElement(leftCorner, "codeSnippetTabLeftCorner");
							addClassToElement(leftCorner, "codeSnippetTabLeftCornerActive");
						}
					}
					if (j == visibleTabs.length-1) {
						addClassToElement(tab, "csLastTab");
						if (hasElementClass(tab, "csActiveTab")) {
							removeClassFromElement(rightCorner, "codeSnippetTabRightCorner");
							addClassToElement(rightCorner, "codeSnippetTabRightCornerActive");
						}
					}
				}
				
				// show/hide code block
				codeCollection = getDivWithClass(snippet, "codeSnippetCodeCollection");
				divs = codeCollection.getElementsByTagName("div");
				for (j=0; j<divs.length; j++) {
					if (hasElementClass(divs[j], "codeSnippetCode")) {
						// it's a code block
						if (langIsNA) {
							showHideTag(divs[j], hasElementClass(divs[j], "codeNA"));
						} else {
							showHideTag(divs[j], hasElementClass(divs[j], languageToShow));
						}
					}
				}
				
			}
		}
	
	} catch (ex) {
	}
}


function showHideTag(tag, visible) {
	try {
		if (visible) {
			tag.style.display = "";
		} else {
			tag.style.display = "none";
		}
	} catch (e) {
	}
}

/***********************************************
 * END LANGUAGE FILTER
 ***********************************************/ 




/***********************************************
 * COPY CODE
 ***********************************************/ 

function CopyCode(item) {
	try {
		// get the visible code block div
		var codeCollection = item.parentNode.parentNode;
		var divs = codeCollection.getElementsByTagName("div");
		var i, shownCode;
		for (i=0; i<divs.length; i++) {
			if (hasElementClass(divs[i], "codeSnippetCode")) {
				// it's a code block
				if (divs[i].style.display != "none") {
					shownCode = divs[i];
					break;
				}
			}
		}

		if (shownCode) {
			// get code and remove <br>
			var code;
			code = shownCode.innerHTML;
			code = code.replace(/<br>/gi, "\n");
			code = code.replace(/<\/td>/gi, "</td>\n");	// syntax highlighter removes \n chars and puts each line in separate <td>
			code = code.trim();	// remove leading spaces which are unwanted in FF 
			// get plain text
			var tmpDiv = document.createElement('div');
			tmpDiv.innerHTML = code;

			if (typeof(tmpDiv.textContent) != "undefined") {
				// standards compliant
				code = tmpDiv.textContent;
			}
			else if (typeof(tmpDiv.innerText) != "undefined") {
				// IE only
				code = tmpDiv.innerText;
			}

			try {
				// works in IE only
				window.clipboardData.setData("Text", code);
			} catch (ex) {
				popCodeWindow(code);
			}
		}
	} catch (e) {
	}
}


function popCodeWindow(code) {
	try {
		var codeWindow =  window.open ("",
			"Copy the selected code",
			"location=0,status=0,toolbar=0,menubar =0,directories=0,resizable=1,scrollbars=1,height=400, width=400");
		codeWindow.document.writeln("<html>");
		codeWindow.document.writeln("<head>");
		codeWindow.document.writeln("<title>Copy the selected code</title>");
		codeWindow.document.writeln("</head>");
		codeWindow.document.writeln("<body bgcolor=\"#FFFFFF\">");
		codeWindow.document.writeln('<pre id="code_text">');
		codeWindow.document.writeln(escapeHTML(code));
		codeWindow.document.writeln("</pre>");
		codeWindow.document.writeln("<scr" + "ipt>");
		// the selectNode function below, converted by http://www.howtocreate.co.uk/tutorials/jsexamples/syntax/prepareInline.html 
		var ftn = "function selectNode (node) {\n\tvar selection, range, doc, win;\n\tif ((doc = node.ownerDocument) && \n\t\t(win = doc.defaultView) && \n\t\ttypeof win.getSelection != \'undefined\' && \n\t\ttypeof doc.createRange != \'undefined\' && \n\t\t(selection = window.getSelection()) && \n\t\ttypeof selection.removeAllRanges != \'undefined\') {\n\t\t\t\n\t\trange = doc.createRange();\n\t\trange.selectNode(node);\n    selection.removeAllRanges();\n    selection.addRange(range);\n\t} else if (document.body && \n\t\t\ttypeof document.body.createTextRange != \'undefined\' && \n\t\t\t(range = document.body.createTextRange())) {\n     \n\t\t \trange.moveToElementText(node);\n     \trange.select();\n  }\n} ";
		codeWindow.document.writeln(ftn);
		codeWindow.document.writeln("selectNode(document.getElementById('code_text'));</scr" + "ipt>");
		codeWindow.document.writeln("</body>");
		codeWindow.document.writeln("</html>");
		codeWindow.document.close();
	} catch (ex) {}
}


function escapeHTML (str) {                                       
	return str.replace(/&/g,"&amp;").                                         
		replace(/>/g,"&gt;").
		replace(/</g,"&lt;").
		replace(/"/g,"&quot;");                                         
}

function selectNode (node) {
	var selection, range, doc, win;
	if ((doc = node.ownerDocument) && 
		(win = doc.defaultView) && 
		typeof win.getSelection != 'undefined' && 
		typeof doc.createRange != 'undefined' && 
		(selection = window.getSelection()) && 
		typeof selection.removeAllRanges != 'undefined') {
			
		range = doc.createRange();
		range.selectNode(node);
    selection.removeAllRanges();
    selection.addRange(range);
	} else if (document.body && 
			typeof document.body.createTextRange != 'undefined' && 
			(range = document.body.createTextRange())) {
     
		 	range.moveToElementText(node);
     	range.select();
  }
} 

/***********************************************
 * END COPY CODE
 ***********************************************/ 


/***********************************************
 * PERSISTENCE
 ***********************************************/ 

/**
 * Sets the cookie value.
 * name - name of the cookie
 * value - value of the cookie
 * [expires] - expiration date of the cookie (defaults to end of current session)
 * [path] - path for which the cookie is valid (defaults to path of calling document)
 * [domain] - domain for which the cookie is valid (defaults to domain of calling document)
 * [secure] - Boolean value indicating if the cookie transmission requires a secure transmission
 * an argument defaults when it is assigned null as a placeholder
 * a null placeholder is not required for trailing omitted arguments
 */
function setCookie(name, value, expires, path, domain, secure) {
  var curCookie = name + "=" + escape(value) +
      ((expires) ? "; expires=" + expires.toGMTString() : "") +
      ((path) ? "; path=" + path : "") +
      ((domain) ? "; domain=" + domain : "") +
      ((secure) ? "; secure" : "");
  document.cookie = curCookie;
}

/**
 * Gets the cookie value.
 * name - name of the desired cookie
 * return string containing value of specified cookie or null if cookie does not exist
 */ 
function getCookie(name) {
  var dc = document.cookie;
  var prefix = name + "=";
  var begin = dc.indexOf("; " + prefix);
  if (begin == -1) {
    begin = dc.indexOf(prefix);
    if (begin != 0) return null;
  } else
    begin += 2;
  var end = document.cookie.indexOf(";", begin);
  if (end == -1)
    end = dc.length;
  return unescape(dc.substring(begin + prefix.length, end));
}


// name - name of the cookie
// [path] - path of the cookie (must be same as path used to create cookie)
// [domain] - domain of the cookie (must be same as domain used to create cookie)
// * path and domain default if assigned null or omitted if no explicit argument proceeds
function deleteCookie(name, path, domain) {
  if (getCookie(name)) {
    document.cookie = name + "=" +
    ((path) ? "; path=" + path : "") +
    ((domain) ? "; domain=" + domain : "") +
    "; expires=Thu, 01-Jan-70 00:00:01 GMT";
  }
}


// fixMoniker is needed to implement userData in a CHM
//
function fixMoniker() {
  var curURL = document.location + ".";
  var pos = curURL.indexOf("mk:@MSITStore");
  if( pos == 0 ) {
    curURL = "ms-its:" + curURL.substring(14,curURL.length-1);
    document.location.replace(curURL);
    return false;
  }
  else { return true; }
}



function saveSetting(name, value) {
	// create an instance of the Date object
	var now = new Date();
	// cookie expires in one year (actually, 365 days)
	// 1000 milliseconds in a second
	now.setTime(now.getTime() + 365 * 24 * 60 * 60 * 1000);
	// convert the value to correct String
	if (value.constructor==Boolean) {
		if (!value) {
			value = "";
		}
	}
	// IE returns wrong document.cookie if the value is empty string
	if (value == "") {
		value = "string:empty"
	}
	
	// we cannot use cookies in CHM, so try to use behaviors if possible
	var headerDiv;	// we can use any particular DIV or other element
	headerDiv = document.getElementById("header");
	if (headerDiv.addBehavior) {
		headerDiv.style.behavior = "url('#default#userData')";
		headerDiv.expires = now.toUTCString();
		headerDiv.setAttribute(name, value);
		// Save the persistence data as "helpSettings".
		headerDiv.save("helpSettings");
	} else {
		// set the new cookie
		setCookie(name, value, now/*, "/"*/);
	}
}


function loadSetting(name, defaultValue) {
	var res;

	// we cannot use cookies in CHM, so try to use behaviors if possible
	var headerDiv;	// we can use any particular DIV or other element
	headerDiv = document.getElementById("header");
	if (headerDiv.addBehavior) {
		headerDiv.style.behavior = "url('#default#userData')";
		headerDiv.load("helpSettings");
		res = headerDiv.getAttribute(name);
	} else {
		// get the cookie
		res = getCookie(name);
	}
	 
	if (res == "string:empty") {
		res = "";
	} 
	if (res == null) {
		res = defaultValue;
	}
	return res;
}


/***********************************************
 * END PERSISTENCE
 ***********************************************/ 




/***********************************************
 * MSDN LINKS FIX
 ***********************************************/ 

/**
 * Fixes all links that should point to MSDN. This
 * applies to CHM documentation. 
 */ 
function fixMsdnLinks()
{
	msdnVersion = getHighestMsdnVersion ();
	//get all MSDN links
	var msdnLinks = new Array();
	var allLinks = document.getElementsByTagName("a");
	var i;
	for (i=0; i<allLinks.length; i++) {
		if (allLinks[i].getAttribute("href")) {
			if (allLinks[i].getAttribute("href").toLowerCase().substr(0,8) == "ms-help:" ) {
				msdnLinks[msdnLinks.length] = allLinks[i]; 
				fixMsdnLink(allLinks[i]);
			}
		}
	}
}


/**
 * Fixes one link that should point to MSDN.
 */ 
function fixMsdnLink(linkElement) {
	var href = linkElement.getAttribute("href");
	if (href) {
		if (href.toLowerCase().substr(0,8) == "ms-help:") {
			var keyword = href.replace(/(.*keyword=\")([^\"]+)(\".*)/g, "$2");
			//fix original MSDN link
			var link;
			switch (msdnVersion) {
				case "9.0":
					link = 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="';
					link = link + escapeKeyword(keyword);
					link = link + '";?index="!DefaultAssociativeIndex"';
					break;
				case "8.0":
					link = 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="';
					link = link + escapeKeyword(keyword);
					link = link + '";?index="!DefaultAssociativeIndex"';
					break;
				default:
					//web
					link = convertMsdnKeywordToMsdn2Link(keyword);
					break
			}
			linkElement.setAttribute("href", link);
		}
	}
}


/**
 * Fixes MSDN keyword for use with redirect.htm in CHM format.
 */ 
function escapeKeyword(keyword) {
  /*
  Keyword can contain "," characters as parameter delimiter. 
  Keyword resolving engine expects TAB where keyword syntax uses ",".
  It also expects NEWLINE where keyword syntax uses ";". While HxLink.htc
  used with <MSHelp:link> in HxS format takes care of it, redirect.htm 
  used in CHM format doesn't. So we must do it.
   */ 
  var res = keyword; 
  res = res.replace(/,/g, "\t");
  res = res.replace(/;/g, "\n");
  res = escape(res);
  return res;
}


/**
 * Converts MSDN2 keyword (cref) to MSDN2 web link
 */ 
function convertMsdnKeywordToMsdn2Link(keyword) {
	var res = keyword;
	
	// remove prefix
	res = res.replace(/^.+:(.*)/g, "$1");
	// remove parameters
	res = res.replace(/(.*)\(.*/g, "$1");
	
	res = res.toLowerCase();
	res = "http://msdn.microsoft.com/en-us/library/" + res + ".aspx";
	return res;
}


// Highest MSDN version installed.
var msdnVersion;


/**
 * Returns highest MSDN version installed.
 * @return The MSDN version found in format 8.0 or 9.0 or web. If none is found,
 *         the "web" is returned.  
 */ 
function getHighestMsdnVersion () {
  var MSDN_9_CSS = "ms-help://MS.VSCC.v90/dv_vscccommon/styles/presentation.css";
	var MSDN_8_CSS = "ms-help://MS.VSCC.v80/dv_vscccommon/local/Classic.css"; 

	if (cssFileExists(MSDN_9_CSS)) {
		return "9.0";
	}
	if (cssFileExists(MSDN_8_CSS)) {
		return "8.0";
	}

  return "web";
}


/**
 * Tests whether specified CSS url exists.
 */ 
function cssFileExists(cssUrl) {
	var sheet = null
	var temporary = false
	res = false;

	// first detect whether this CSS is already used in this document (it should be)
  try {
    var i;
    for (i=0; i<document.styleSheets.length; i++) {
      if (document.styleSheets[i].href.toLowerCase() == cssUrl.toLowerCase()) {
        sheet = document.styleSheets[i];
        break;
      }    
    }
  } catch (ex) {
  }
  
	// now check, if the sheet really contains any rules - the CSS file exists
	try {
		if (sheet.rules) {
			// IE
			if (sheet.rules.length > 0) {
				res = true
			}
		} else if (sheet.cssRules) {
			// FF
			if (sheet.cssRules.length > 0) {
				res = true
			}
		}
	} catch (ex) {
	}

	return res;
}

/***********************************************
 * END MSDN LINKS FIX
 ***********************************************/ 
