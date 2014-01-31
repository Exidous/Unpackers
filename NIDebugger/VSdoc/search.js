function stripSpaces(text) {
//    return text.replace(/^\W+/,'').replace(/\W+$/,'');
    return text.split(" ").join(" ");
}

function get_bitmap(setbit, size) {
    var map = {};
    for(var i=0;i<size;++i) map[i]=0;
    map[setbit]=1;
    return map;
}

function do_search() {
    // hide all documents
    for(var i=0;i<nurls;i++) document.getElementById("s_"+i).style.display="none";
    document.getElementById("content").style.display="none";
    document.getElementById("no_results").style.display="none";
    document.getElementById("some_results").style.display="none";
        
    // extract all search terms
    var term = stripSpaces(document.getElementById("q").value.toLowerCase());
    qterms = [];
    var chunks = term.split(" ");
    for(var i in chunks)
        if(chunks[i]) qterms[qterms.length] = chunks[i];

    var candidates = {};
    for(var i in qterms) {
        var term = qterms[i];
				/*   
        if(index[term]!=undefined) {
        	// whole words
            for(var docid in index[term]) {
                if(candidates[index[term][docid]]==undefined)
                    candidates[index[term][docid]]=get_bitmap(i,qterms.length);
                else
                    candidates[index[term][docid]][i]=1;
            }        
        }
        else {
        */
        	// parts of words
            for(var keyid in index) {
                if(keyid.indexOf(term)>=0) {
                    for(var docid in index[keyid]) {
                        if(candidates[index[keyid][docid]]==undefined)
                            candidates[index[keyid][docid]]=get_bitmap(i,qterms.length);
                        else 
                            candidates[index[keyid][docid]][i]=1;
                    }
                }
            }
        //}
    }
    
    
    var somethingFound = false;
    for(var key in candidates) {
        var on=1;
        for(var i in qterms) on = on && candidates[key][i];
        if(on) {
					document.getElementById("s_"+key).style.display="list-item";
					somethingFound = true;
				}
    }
    
    if (somethingFound){
    	document.getElementById("content").style.display="block";
    	document.getElementById("some_results").style.display="block";
		} else {
    	document.getElementById("no_results").style.display="block";
    }
}

function do_highlight() {
	try {
    colors=['yellow','lightgreen','gold','orange','magenta','lightblue'];
    var searchFrame = top.frames["vbdocswitch"].frames["main-iframe"]
    if(searchFrame.qterms!=undefined) {
        for(var i in searchFrame.qterms) {
            parent.frames['vbdocright'].document.body.innerHTML = doHighlight(parent.frames['vbdocright'].document.body.innerHTML,searchFrame.qterms[i],colors[i%colors.length]);
						parent.frames['vbdocright'].init(); 
        }
    }
  } catch (ex) {}
}

// from http://www.nsftools.com/misc/SearchAndHighlight.htm

/*
 * This is the function that actually highlights a text string by
 * adding HTML tags before and after all occurrences of the search
 * term. You can pass your own tags if you'd like, or if the
 * highlightStartTag or highlightEndTag parameters are omitted or
 * are empty strings then the default <font> tags will be used.
 */
function doHighlight(bodyText, searchTerm, color, highlightStartTag, highlightEndTag) 
{
  // the highlightStartTag and highlightEndTag parameters are optional
  if ((!highlightStartTag) || (!highlightEndTag)) {
    highlightStartTag = "<font style='background-color:"+color+";'>";
    highlightEndTag = "</font>";
  }
  
  // find all occurences of the search term in the given text,
  // and add some "highlight" tags to them (we're not using a
  // regular expression search, because we want to filter out
  // matches that occur within HTML tags and script blocks, so
  // we have to do a little extra validation)
  var newText = "";
  var i = -1;
  var lcSearchTerm = searchTerm.toLowerCase();
  var lcBodyText = bodyText.toLowerCase();
    
  while (bodyText.length > 0) {
    i = lcBodyText.indexOf(lcSearchTerm, i+1);
    if (i < 0) {
      newText += bodyText;
      bodyText = "";
    } else {
      // skip anything inside an HTML tag
      if (bodyText.lastIndexOf(">", i) >= bodyText.lastIndexOf("<", i)) {
        // skip anything inside a <script> block
        if (lcBodyText.lastIndexOf("/script>", i) >= lcBodyText.lastIndexOf("<script", i)) {
          newText += bodyText.substring(0, i) + highlightStartTag + bodyText.substr(i, searchTerm.length) + highlightEndTag;
          bodyText = bodyText.substr(i + searchTerm.length);
          lcBodyText = bodyText.toLowerCase();
          i = -1;
        }
      }
    }
  }
  
  return newText;
}