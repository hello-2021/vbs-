dim password
do
password=inputbox("想玩QQ？先叫爸爸！","腾讯QQ")
if password="爸爸"then
msgbox("哎，真乖，给你玩吧！")
set WShShell=createobject("wscript.shell")
WshShell.Exec("C:\Program Files (x86)\Tencent\QQ\Bin\QQScLauncher.exe")
exit do
end if
loop