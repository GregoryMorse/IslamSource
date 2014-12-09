// JavaScript Document

function getBaseURL () {
   return location.protocol + '//' + location.hostname + 
      (location.port && ':' + location.port) + '/';
}

(function() {
    tinymce.create('tinymce.plugins.is_button', {
        init : function(ed, url) {
            ed.addButton('is_button', {
                title : 'IslamSource',
                image : url+'/render.php?Size=32&Char=22&Font=AGAIslamicPhrases',
      			onclick : function() {
      				ed.selection.setContent('{}');
                }
            });
        },
        createControl : function(n, cm) {
            return null;
        },
    });
    tinymce.PluginManager.add('is_button', tinymce.plugins.is_button);
})();