# IslamSource
Contributors: GregoryMorse
Donate link: http://www.islamSource.info/
Tags: Islam,Quran,calligraphy,phrases,verses
Requires at least: 3.0.1
Tested up to: 4.3.1
Stable tag: trunk
License: GPLv2 or later
License URI: http://www.gnu.org/licenses/gpl-2.0.html

Islam Source Plugin to insert Quranic verses and calligraphy of Islamic phrases easily and simply.

## Description

Islam Source Quranic Verse and Islamic Phrase Plugin - Allows for Quranic chapters, verses even specified down through the word to be inserted easily by using formats {a:b:c-x:y:z} where b, c, x, y and z are optional depending on if a chapter, verse or particular word of a verse is desired or a range is desired so it could be in forms {a:b-y} or {a:b-x:y} such as the opening chapter which could be specified as {1:1-7}.  The Arabic is automatically displayed when posts are viewed.  It also allows for various calligraphy and Unicode Islamic words or phrases to be easily inserted through a button on the visual editor which are displayed when the posts are later viewed.

## Installation

1. Automated installation through the plugin menu, at which point the icon will appear in the Visual editor and any matching tags will be shown.
2. Manual installation: Difficulties due to the large size of the plugin for the metadata, translation and Quran resources may require adjusting the web server at least temporarily to upload the plugin: PHP in php.ini: post_max_size = 128M and upload_max_filesize = 128M, IIS in applicationhost.config or web.config: configuration -> system.webServer -> security -> requestFiltering -> add or change `<requestLimits maxAllowedContentLength="134217728" />`
3. Extract `http://www.islamsource.info/IslamSource.zip` to the `/wp-content/plugins/` directory.
4. Activate the plugin through the 'Plugins' menu in WordPress.

## Frequently Asked Questions

What parameters can be set?

    {a:b:c-x:y:z;parameter=value;parameter}

e.g. {1:1-2;W4W;TranslitStyle=RuleBased;TranslitScheme=PlainRoman;Translation=en.pickthall}

"Header"
Turn on the Ruku, Surah, Ayah header display.

"W4W"
Enable Word-for-Word mode.

"W4WNum"
Enable Word-for-Word mode showing the number of each word in the verse.

"NoArabic"
Disable the Arabic as well as transliteration.

"TranslitStyle"="Literal"|"RuleBased"|"LearningMode"
Literal is a character by character transliteration, rule based applies tajweed rules and learning mode applies tajweed rules showing options for stopping/continuing

"TranslitScheme=<name from list below>"
Name Transliteration 
ExtendedBuckwalter  Extended Buckwalter  
IPAValues  IPA Values  
RomanTranslit  Roman Transliteration  
HansWeir  Hans Weir  
HansWeir4thEdition  Hans Weir 4th Edition  
PlainRoman  Plain Roman  
InternetArabic  Internet Arabic 

"Translation="
default to en.sahih  Saheeh International (Saheeh International)
"Translation=<name from list below>"
Name Translation 
en-maulana_ali  Maulana Muhammad Ali (Maulana Muhammad Ali)  
en.asad  Muhammad Asad (Asad)  
sq.nahi  Hasan Efendi Nahi (Efendi Nahi)  
sq.mehdiu  Feti Mehdiu (Feti Mehdiu)  
sq.ahmeti  Sherif Ahmeti (Sherif Ahmeti)  
ber.mensur  Ramdane At Mansour (At Mensur)  
ar.jalalayn  Jalal ad-Din al-Mahalli and Jalal ad-Din as-Suyuti (تفسير الجلالين)  
ar.muyassar  King Fahad Quran Complex (تفسير المیسر)  
am.sadiq  Muhammed Sadiq and Muhammed Sani Habib (ሳዲቅ & ሳኒ ሐቢብ)  
az.mammadaliyev  Vasim Mammadaliyev and Ziya Bunyadov (Məmmədəliyev & Bünyadov)  
az.musayev  Alikhan Musayev (Musayev)  
bn.hoque  Zohurul Hoque (জহুরুল হক)  
bn.bengali  Muhiuddin Khan (মুহিউদ্দীন খান)  
bs.korkut  Besim Korkut (Korkut)  
bs.mlivo  Mustafa Mlivo (Mlivo)  
bg.theophanov  Tzvetan Theophanov (Теофанов)  
zh.jian  Ma Jian (Ma Jian)  
zh.majian  Ma Jian (Ma Jian (Traditional))  
cs.hrbek  Preklad I. Hrbek (Hrbek)  
cs.nykl  A. R. Nykl (Nykl)  
dv.divehi  Office of the President of Maldives (ދިވެހި)  
nl.keyzer  Salomo Keyzer (Keyzer)  
nl.leemhuis  Fred Leemhuis (Leemhuis)  
nl.siregar  Sofian S. Siregar (Siregar)  
en.ahmedali  Ahmed Ali (Ahmed Ali)  
en.ahmedraza  Ahmed Raza Khan (Ahmed Raza Khan)  
en.arberry  A. J. Arberry (Arberry)  
en.daryabadi  Abdul Majid Daryabadi (Daryabadi)  
en.hilali  Muhammad Taqi-ud-Din al-Hilali and Muhammad Muhsin Khan (Hilali & Khan)  
en.itani  Talal Itani (Itani)  
en.maududi  Abul Ala Maududi (Maududi)  
en.pickthall  Mohammed Marmaduke William Pickthall (Pickthall)  
en.qarai  Ali Quli Qarai (Qarai)  
en.qaribullah  Hasan al-Fatih Qaribullah and Ahmad Darwish (Qaribullah & Darwish)  
en.sahih  Saheeh International (Saheeh International)  
en.sarwar  Muhammad Sarwar (Sarwar)  
en.shakir  Mohammad Habib Shakir (Shakir)  
en.transliteration  English Transliteration (Transliteration)  
en.wahiduddin  Wahiduddin Khan (Wahiduddin Khan)  
en.yusufali  Abdullah Yusuf Ali (Yusuf Ali)  
fr.hamidullah  Muhammad Hamidullah (Hamidullah)  
de.aburida  Abu Rida Muhammad ibn Ahmad ibn Rassoul (Abu Rida)  
de.bubenheim  A. S. F. Bubenheim and N. Elyas (Bubenheim & Elyas)  
de.khoury  Adel Theodor Khoury (Khoury)  
de.zaidan  Amir Zaidan (Zaidan)  
ha.gumi  Abubakar Mahmoud Gumi (Gumi)  
hi.farooq  Muhammad Farooq Khan and Muhammad Ahmed (फ़ारूक़ ख़ान & अहमद)  
hi.hindi  Suhel Farooq Khan and Saifur Rahman Nadwi (फ़ारूक़ ख़ान & नदवी)  
id.indonesian  Indonesian Ministry of Religious Affairs (Bahasa Indonesia)  
id.muntakhab  Muhammad Quraish Shihab et al. (Quraish Shihab)  
id.jalalayn  Jalal ad-Din al-Mahalli and Jalal ad-Din as-Suyuti (Tafsir Jalalayn)  
it.piccardo  Hamza Roberto Piccardo (Piccardo)  
ja.japanese  Unknown (Japanese)  
ko.korean  Unknown (Korean)  
ku.asan  Burhan Muhammad-Amin (ته‌فسیری ئاسان)  
ms.basmeih  Abdullah Muhammad Basmeih (Basmeih)  
ml.abdulhameed  Cheriyamundam Abdul Hameed and Kunhi Mohammed Parappoor (അബ്ദുല്‍ ഹമീദ് & പറപ്പൂര്‍)  
ml.karakunnu  Muhammad Karakunnu and Vanidas Elayavoor (കാരകുന്ന് & എളയാവൂര്)  
no.berg  Einar Berg (Einar Berg)  
fa.ghomshei  Mahdi Elahi Ghomshei (الهی قمشه‌ای)  
fa.ansarian  Hussain Ansarian (انصاریان)  
fa.ayati  AbdolMohammad Ayati (آیتی)  
fa.bahrampour  Abolfazl Bahrampour (بهرام پور)  
fa.khorramdel  Mostafa Khorramdel (خرمدل)  
fa.khorramshahi  Baha'oddin Khorramshahi (خرمشاهی)  
fa.sadeqi  Mohammad Sadeqi Tehrani (صادقی تهرانی)  
fa.fooladvand  Mohammad Mahdi Fooladvand (فولادوند)  
fa.mojtabavi  Sayyed Jalaloddin Mojtabavi (مجتبوی)  
fa.moezzi  Mohammad Kazem Moezzi (معزی)  
fa.makarem  Naser Makarem Shirazi (مکارم شیرازی)  
pl.bielawskiego  Józefa Bielawskiego (Bielawskiego)  
pt.elhayek  Samir El-Hayek (El-Hayek)  
ro.grigore  George Grigore (Grigore)  
ru.abuadel  Abu Adel (Абу Адель)  
ru.muntahab  Ministry of Awqaf, Egypt (Аль-Мунтахаб)  
ru.krachkovsky  Ignaty Yulianovich Krachkovsky (Крачковский)  
ru.kuliev  Elmir Kuliev (Кулиев)  
ru.kuliev-alsaadi  Elmir Kuliev (with Abd ar-Rahman as-Saadi's commentaries) (Кулиевас-Саади)  
ru.osmanov  Magomed-Nuri Osmanovich Osmanov (Османов)  
ru.porokhova  V. Porokhova (Порохова)  
ru.sablukov  Gordy Semyonovich Sablukov (Саблуков)  
sd.amroti  Taj Mehmood Amroti (امروٽي)  
so.abduh  Mahmud Muhammad Abduh (Abduh)  
es.asad  Muhammad Asad - Abdurrasak Pérez (Asad)  
es.bornez  Raúl González Bórnez (Bornez)  
es.cortes  Julio Cortes (Cortes)  
es.garcia  Muhammad Isa García (Garcia)  
sw.barwani  Ali Muhsin Al-Barwani (Al-Barwani)  
sv.bernstrom  Knut Bernström (Bernström)  
tg.ayati  AbdolMohammad Ayati (Оятӣ)  
ta.tamil  Jan Turst Foundation (ஜான் டிரஸ்ட்)  
tt.nugman  Yakub Ibn Nugman (Yakub Ibn Nugman)  
th.thai  King Fahad Quran Complex (ภาษาไทย)  
tr.golpinarli  Abdulbaki Golpinarli (Abdulbakî Gölpınarlı)  
tr.bulac  Alİ Bulaç (Alİ Bulaç)  
tr.transliteration  Muhammet Abay (Çeviriyazı)  
tr.diyanet  Diyanet Isleri (Diyanet İşleri)  
tr.vakfi  Diyanet Vakfi (Diyanet Vakfı)  
tr.yuksel  Edip Yüksel (Edip Yüksel)  
tr.yazir  Elmalili Hamdi Yazir (Elmalılı Hamdi Yazır)  
tr.ozturk  Yasar Nuri Ozturk (Öztürk)  
tr.yildirim  Suat Yildirim (Suat Yıldırım)  
tr.ates  Suleyman Ates (Süleyman Ateş)  
ur.maududi  Abul A'ala Maududi (ابوالاعلی مودودی)  
ur.kanzuliman  Ahmed Raza Khan (احمد رضا خان)  
ur.ahmedali  Ahmed Ali (احمد علی)  
ur.jalandhry  Fateh Muhammad Jalandhry (جالندہری)  
ur.qadri  Tahir ul Qadri (طاہر القادری)  
ur.jawadi  Syed Zeeshan Haider Jawadi (علامہ جوادی)  
ur.junagarhi  Muhammad Junagarhi (محمد جوناگڑھی)  
ur.najafi  Ayatollah Muhammad Hussain Najafi (محمد حسین نجفی)  
ug.saleh  Muhammad Saleh (محمد صالح)  
uz.sodik  Muhammad Sodik Muhammad Yusuf (Мухаммад Содик)  

== Screenshots ==

1. The visual editor showing the toolbar, popup and style of tags.
2. The rendered result when viewing the post.

== Changelog ==

= 1.6 = 
Updated metadata system for latest changes involving learning mode and contiguity and some efficiency enhancements.

= 1.3 =
Better organization, error checking, fixes and ability to add Arabic text.

= 1.2 =
Speed optimizations.

= 1.1 =
Fixes and translations.

= 1.0 =
Initial version.

# Upgrade Notice

None
