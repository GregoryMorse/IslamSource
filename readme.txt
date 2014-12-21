=== IslamSource ===
Contributors: GregoryMorse
Donate link: http://www.islamSource.info/
Tags: Islam,Quran,calligraphy,phrases,verses
Requires at least: 3.0.1
Tested up to: 4.1.0
Stable tag: trunk
License: GPLv2 or later
License URI: http://www.gnu.org/licenses/gpl-2.0.html

Islam Source Plugin to insert Quranic verses and calligraphy of Islamic phrases easily and simply.

== Description ==

Islam Source Quranic Verse and Islamic Phrase Plugin - Allows for Quranic chapters, verses even specified down through the word to be inserted easily by using formats {a:b:c-x:y:z} where b, c, x, y and z are optional depending on if a chapter, verse or particular word of a verse is desired or a range is desired so it could be in forms {a:b-y} or {a:b-x:y} such as the opening chapter which could be specified as {1:1-7}.  The Arabic is automatically displayed when posts are viewed.  It also allows for various calligraphy and Unicode Islamic words or phrases to be easily inserted through a button on the visual editor which are displayed when the posts are later viewed.

== Installation ==

1. Automated installation through the plugin menu, at which point the icon will appear in the Visual editor and any matching tags will be shown.
2. Manual installation: Difficulties due to the large size of the plugin for the metadata, translation and Quran resources may require adjusting the web server at least temporarily to upload the plugin: PHP in php.ini: post_max_size = 128M and upload_max_filesize = 128M, IIS in applicationhost.config or web.config: configuration -> system.webServer -> security -> requestFiltering -> add or change <requestLimits maxAllowedContentLength="134217728" />
3. Extract `http://www.islamsource.info/IslamSource.zip` to the `/wp-content/plugins/` directory.
4. Activate the plugin through the 'Plugins' menu in WordPress.

== Frequently Asked Questions ==

None

== Screenshots ==

1. The visual editor showing the toolbar, popup and style of tags.
2. The rendered result when viewing the post.

== Changelog ==

= 1.1 =
Fixes and translations.

= 1.0 =
Initial version.

== Upgrade Notice ==

None