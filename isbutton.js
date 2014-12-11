// JavaScript Document

function getBaseURL () {
   return location.protocol + '//' + location.hostname + 
      (location.port && ':' + location.port) + '/';
}

(function() {
    tinymce.create('tinymce.plugins.is_button', {
        init : function(ed, url) {
			var xmlhttp = new XMLHttpRequest();
			xmlhttp.open("GET", url + '/render.php?Cat=', false);
			xmlhttp.send();
		    var data = JSON.parse(xmlhttp.responseText);
		    for (var count = 0; count < data.length; count++) {
		    	for (var subcount = 0; subcount < data[count].menu.length; subcount++) {
		    		if (data[count].menu[subcount] !== undefined && data[count].menu[subcount] !== null) {
			    		var val = [], getClickFunc = function(paste, ch) {
		    					return function(e) {
				      				ed.selection.setContent((ch ? '' : '{') + paste + (ch ? '' : '}'));
		      					};
			    			}, getMenuClickFunc = function(v) {
					    		return function () {
				      				ed.windowManager.open({
				      					scrollbars: true,
				      					title: 'Choose Item(s)',
				      					body: v
				      				});
				      			};
					    	}, getOnPostRender = function(u, ch) { return function () { this.getEl().style.height = "auto"; this.getEl().firstChild.innerHTML = ch ? "<div style=\"font-size: 100px; display: inline-block;\">" + u + "</div>" : "<img src=\"" + u + "\" style=\"width=" + this.getEl().style.width + ";height=auto;\" onload=\"" + "\">"; }; };
				    	for (var pancount = 0; pancount < data[count].menu[subcount].values.length; pancount++) {
				    		if (data[count].menu[subcount].values[pancount] !== null) {
				    			if (data[count].menu[subcount].values[pancount].font !== "") {
						    		val.push({type: 'button', name: 'category' + pancount, onclick: getClickFunc(data[count].menu[subcount].values[pancount].value + ";" + 'Font=' + data[count].menu[subcount].values[pancount].font + ';Char=' + data[count].menu[subcount].values[pancount].char, false),
						    			onPostRender: getOnPostRender(url + '/render.php?Size=100&Font=' + data[count].menu[subcount].values[pancount].font + '&Char=' + data[count].menu[subcount].values[pancount].char, false)
						    			});
						    	} else {
						    		val.push({type: 'button', name: 'category' + pancount, onclick: getClickFunc('&#x' + data[count].menu[subcount].values[pancount].char + ';', true), style: "text-align: center;",
						    			onPostRender: getOnPostRender(String.fromCharCode(parseInt(data[count].menu[subcount].values[pancount].char, 16)), true)
						    			});
						    	}
		                    }
				    	}
				    	data[count].menu[subcount].onclick = getMenuClickFunc(val);
				    }
			    }
		    }
            ed.addButton('is_button', {
                type: 'menubutton',
                title: 'Islam Source',
                image: url+'/render.php?Size=32&Char=22&Font=AGAIslamicPhrases',
                style: 'background-image: url(' + url+'/render.php?Size=32&Char=22&Font=AGAIslamicPhrases); background-size: contain; background-repeat: no-repeat;',
                icons: false,
                //text: 'Islam Source',
                menu: data
			});
        },
        createControl : function(n, cm) {
            return null;
        },
    });
    tinymce.PluginManager.add('is_button', tinymce.plugins.is_button);
})();