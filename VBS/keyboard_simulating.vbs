set mySendKeys = CreateObject("WScript.Shell")
mySendKeys.SendKeys "Hello"
mySendKeys.SendKeys "{DOWN}"
mySendKeys.SendKeys "{EMTER}"
set mySendKeys = NOTHING