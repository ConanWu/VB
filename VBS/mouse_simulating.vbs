systemutil.Run  "IExplore.exe", "www.baidu.com"
Setting.WebPackage("ReplayType")=2   '鼠标跟踪模式
'配置使用浏览器事件或鼠标运行鼠标操作的'方式.1 - 使用浏览器事件运行鼠标操作,2 - 使用鼠标运行鼠标操作。
Browser("百度一下，你就知道").Page("百度一下，你就知道").Link("关于百度").FireEvent("onmouseover") '鼠标移动到指定对象上面
Setting.WebPackage("ReplayType")=1   '还原鼠标默认模式

'Fireevent事件:单击（onclick）、双击(onDblClick)、光标聚集（OnBlur）、onchange、onfocus、onmousedown、onmouseup、onmouseover、onmouseout、onsubmit、onreset、onpropertychange
