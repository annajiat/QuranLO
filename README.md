# QuranLO
LibreOffice extension to add Qur'an text to a document. It allows you to select a complete 
surah or a range of ayats of a surah.

The standard font is selected from the LibreOffice Basic Fonts (CTL) setting. 
It uses the Default font and its Fontsize. The Font and Fontsize can be overridden by setting 
it on the selection dialog. The Font Selection shows only fonts that support the arabic ...

A fonts that gives very good results is: 
**Scheherazade**. This you can find at <https://software.sil.org/scheherazade/>.
It allows you to fine tune the font by so called smart features. You can even generate a font with the smart features set to your liking. 

**Al Qalam Quran Majeed**, **Al Qalam Quran Majeed 1**, **Al Qalam Quran Majeed 2**. 

Other Arabic fonts give mixed results. Some don't have the parenthesis or even the standard Arabic numbers. I tried to mitigate for that by providing substitutions.

To build the extension I used the LibreOffice Eclipse plugin for extension development: 
<https://libreoffice.github.io/loeclipse/>. It also provides a starter project that you can use as an example. 

The Qur'an text and its translations were provided by <https://tanzil.net>:  

  Tanzil Quran Text (Uthmani, version 1.0.2)  
  Copyright (C) 2008-2010 Tanzil.net..  
  License: Creative Commons Attribution 3.0  
