Set WshShell= WScript.CreateObject("WScript.shell")
WshShell.AppActivate"114514"
dim times
dim text
times=inputbox("输入次数","输入轰炸次数")
text=inputbox("输入内容","输入轰炸内容")
WshShell.Run "cmd.exe /c echo " & text & "| clip",0,False
msgbox "按确定后倒计时1秒开始",,"提示"
WScript.sleep 1000
for i=1 to times
WScript.sleep 200
WshShell.SendKeys "^v"
WshShell.SendKeys "{ENTER}"
Next
WScript.sleep 100
msgbox "轰炸完成",,"提示"