<h1>Google TTS for Windows on VBScript</h1>

Generates a mp3 file from the text using Google TTS, for any supported Google language

You can run the script with a command-line option:

<b>/lang:uk</b>
<br />
<b>/utf8:false</b>
<br />
<b>/clipboard</b>

uk - generation language
<br />
utf8 - default UTF8 Encode (True, False)
<br />
clipboard - read text from clipboard

You can download MADPlay from <a href="http://www.softpedia.com/get/Multimedia/Audio/Other-AUDIO-Tools/?utm_source=spd&utm_campaign=postdl_redir">here...</a>
(The MADPlay application was developed to be a small, easy to use command line MP3 player / decoder that will allow you to decode and play MPEG audio file.)

Example of use In the bat file: (use CP866)
<code><pre>
Set f_time=%time:~0,5%
Set h_time=%time:~0,2%

Rem EQU - дорiвнює
Rem NEQ - не дорiвнює
Rem LSS - меньше
Rem LEQ - меньше або дорiвнює
Rem GTR - бiльше
Rem GEQ - бiльше або дорiвнює

Rem доброї ночi
If %h_time% GEQ 00 Set greet=доброї ночi
Rem доброго ранку
If %h_time% GEQ 05 Set greet=доброго ранку
Rem доброго дня
If %h_time% GEQ 10 Set greet=доброго дня
Rem доргого вечора
If %h_time% GEQ 19 Set greet=доброго вечора
Rem доброї ночi
If %h_time% GEQ 23 Set greet=доброї ночi

start /MIN wscript google_sapi.vbs /lang:uk Юрiю %greet%. Система Успiшно завантажена. Пiдключення до Iнтернету встановлено. Поточний час %f_time%
</pre></code>

<b>Yurii Radio - 2017</b>
