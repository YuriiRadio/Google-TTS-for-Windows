<h1>Google SAPI for Windows on VBScript</h1>

Generates a mp3 file from the text using Google SAPI.

You can run the script with a command-line option:

<b>/lang:uk</b>
<br />
<b>/utf8:false</b>

uk - generation language
<br />
utf8 - default UTF8 Encode (True, False)

Example of use In the bat file:
<code>
Set f_time=%time:~0,5%
Set h_time=%time:~0,2%

Rem EQU - дорiвнює
Rem NEQ - не дорiвнює
Rem LSS - меноше
Rem LEQ - меньше або дор?внює
Rem GTR - бiльше
Rem GEQ - бiльше або дор?внює

Rem доброї ноч?
If %h_time% GEQ 00 Set greet=доброї ночi
Rem доброго ранку
If %h_time% GEQ 05 Set greet=доброго ранку
Rem доброго дня
If %h_time% GEQ 10 Set greet=доброго дня
Rem доргого вечора
If %h_time% GEQ 19 Set greet=доброго вечора
Rem доброї ноч?
If %h_time% GEQ 23 Set greet=доброї ночi

start /MIN wscript google_sapi.vbs /lang:uk Юрiю %greet%. Система Успiшно завантажена. Пiдключення до Iнтернету встановлено. Поточний час %f_time%
</code>

<b>Yurii Radio - 2017</b>
