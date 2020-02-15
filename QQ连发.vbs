On Error Resume Next 
Dim wsh,ye,who,num
set wsh=createobject("wscript.shell") 
who=inputbox("将剪贴板的最新项连法10次到QQ的哪个聊天","QQ连发","你的好友")
for i = 1 to 10 
    wscript.sleep 700 
    wsh.AppActivate(who) 
    wsh.sendKeys "^v" 
    wsh.sendKeys i 
    wsh.sendKeys "%s" 
next 
wscript.quit