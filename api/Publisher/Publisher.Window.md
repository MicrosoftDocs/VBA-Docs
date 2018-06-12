---
title: Window Object (Publisher)
keywords: vbapb10.chm327679
f1_keywords:
- vbapb10.chm327679
ms.prod: publisher
api_name:
- Publisher.Window
ms.assetid: 342d77cd-5556-6ac3-a828-b1b60380f910
ms.date: 06/08/2017
---


# Window Object (Publisher)

Represents a window. Many publication characteristics, such as scroll bars and rulers, are actually properties of the window.
 


## Example

Use the  **[ActiveWindow](Publisher.Application.ActiveWindow.md)** property to return a **Window** object. The following example maximizes the active window.
 

 

```
Sub MaximizeWindow 
 ActiveWindow.WindowState = pbWindowStateMaximize 
End Sub
```

Use the  **[Caption](Publisher.Window.Caption.md)** property to return the file and application names of the active window. The following example displays a message with the file name and Microsoft Publisher application name.
 

 



```
Sub ShowFileApNames 
 MsgBox Windows(1).Caption 
End Sub
```


## Methods



|**Name**|
|:-----|
|[Activate](Publisher.Window.Activate.md)|
|[Move](Publisher.Window.Move.md)|
|[Resize](Publisher.Window.Resize.md)|

## Properties



|**Name**|
|:-----|
|[Application](Publisher.Window.Application.md)|
|[Caption](Publisher.Window.Caption.md)|
|[Height](Publisher.Window.Height.md)|
|[Hwnd](Publisher.Window.Hwnd.md)|
|[Left](Publisher.Window.Left.md)|
|[Parent](Publisher.Window.Parent.md)|
|[Top](Publisher.Window.Top.md)|
|[Visible](Publisher.Window.Visible.md)|
|[Width](Publisher.Window.Width.md)|
|[WindowState](window-windowstate-property-publisher.md)|

