<?php
//GD library and FreeType library
$fonts = array("AGAIslamicPhrases", "AGAArabesque", "Shia", "IslamicLogo", "KFGQPCArabicSymbols01", "Quranic", "Tulth", "Farsi", "Asmaul-Husna", "Asmaul-Husna_2");
$fontfiles = array("AGA_Islamic_Phrases.TTF", "aga-arabesque.ttf", "SHIA.TTF", "islamic.ttf", "Symbols1_Ver02.otf", "Quranic.ttf", "Tulth.ttf", "Farsi.ttf", "Asmaul-Husna_1.ttf", "Asmaul-Husna_2.ttf");
header('Content-Type: image/png');
$font = '\\files\\' . $fontfiles[0];
$text = chr(0x22);
$bbox = imagettfbbox(32, 0, $font, $text);
$im = imagecreatetruecolor(ceil($bbox[4] - $bbox[0]), ceil($bbox[1] - $bbox[5]));
$white = imagecolorallocate($im, 255, 255, 255);
$black = imagecolorallocate($im, 0, 0, 0);
imagefilledrectangle($im, 0, 0, ceil($bbox[4] - $bbox[0]), ceil($bbox[1] - $bbox[5]), $white);
imagettftext($im, 32, 0, 0, $bbox[1], $black, $font, $text);
imagepng($im);
imagedestroy($im);
?>