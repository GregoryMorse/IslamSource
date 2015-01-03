// JavaScript Document

function getBaseURL () {
   return location.protocol + '//' + location.hostname + 
      (location.port && ':' + location.port) + '/';
}

(function() {
    tinymce.create('tinymce.plugins.is_button', {
        init : function(ed, url) {
			var xmlhttp = new XMLHttpRequest();
			xmlhttp.open("GET", url + '/IslamSourceWP.php?Cat=', false);
			xmlhttp.send();
		    var data = JSON.parse(xmlhttp.responseText);
		    for (var count = 0; count < data.length; count++) {
		    	for (var subcount = 0; subcount < data[count].menu.length; subcount++) {
		    		if (data[count].menu[subcount] !== undefined && data[count].menu[subcount] !== null) {
			    		var val = [], getClickFunc = function(b, paste, ch) {
		    					return function(e) {
				      				ed.selection.setContent((ch ? '' : '{') + paste + (ch ? '' : ";Size=" + (parseInt(ed.queryCommandValue('FontSize')) || parseInt(tinyMCE.DOM.getStyle(ed.selection.getNode(), 'fontSize', true), 10)) + '}'));
		      					};
			    			}, getMenuClickFunc = function(v) {
					    		return function () {
				      				ed.windowManager.open({
				      					width: 362,
				      					height: 300,
				      					scrollbars: true,
				      					style: 'overflow: scroll;',
				      					title: 'Choose Item(s)',
				      					buttons: [],
				      					body: v
				      				});
				      			};
					    	}, getOnPostRender = function(u, ch) { return function () { var el = document.createElement(ch ? "span" : "img"); var size = parseInt(ed.queryCommandValue('FontSize')) || parseInt(tinyMCE.DOM.getStyle(ed.selection.getNode(), 'fontSize', true), 10); var fontName = ed.queryCommandValue('FontName') || tinyMCE.DOM.getStyle(ed.selection.getNode(), 'fontFamily', true); if (ch) { el.style.fontSize = size + "px"; el.style.display = "inline-block"; el.innerText = u; } else { el.style.maxWidth = "280px"; el.style.width = "auto"; el.style.height = "auto"; var f = function() { if (!this.complete) { window.setTimeout(f.bind(this), 1000); } else { this.parentNode.style.maxWidth = '300px'; this.style.maxWidth = '280px'; this.style.width = 'auto'; } }; el.onload = f; el.src = u + "&Size=" + size; } this.getEl().style.maxWidth = '300px'; this.getEl().style.height = "auto"; this.getEl().firstChild.appendChild(el); if(!ch) el.onload(); }; };
				    	for (var pancount = 0; pancount < data[count].menu[subcount].values.length; pancount++) {
				    		if (data[count].menu[subcount].values[pancount] !== null) {
				    			if (data[count].menu[subcount].values[pancount].font !== "") {
						    		val.push({type: 'button', name: 'category' + pancount, onclick: getClickFunc(this, data[count].menu[subcount].values[pancount].value + ";" + 'Font=' + data[count].menu[subcount].values[pancount].font + ';Char=' + data[count].menu[subcount].values[pancount].char, false),
						    			onPostRender: getOnPostRender(url + '/IslamSourceWP.php?Font=' + data[count].menu[subcount].values[pancount].font + '&Char=' + data[count].menu[subcount].values[pancount].char, false)
						    			});
						    	} else {
						    		var st = '';
						    		for (var chcount = 0; chcount < data[count].menu[subcount].values[pancount].char.length; chcount++) { st += '&#x' + data[count].menu[subcount].values[pancount].char.charCodeAt(chcount).toString(16) + ';'; }
						    		val.push({type: 'button', name: 'category' + pancount, onclick: getClickFunc(this, st, true), style: "text-align: center;",
						    			onPostRender: getOnPostRender(data[count].menu[subcount].values[pancount].char, true)
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
                image: url + '/IslamSourceWP.php?Size=32&Char=22&Font=AGAIslamicPhrases',
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