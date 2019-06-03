---
title: Window object (Publisher)
keywords: vbapb10.chm327679
f1_keywords:
- vbapb10.chm327679
ms.prod: publisher
api_name:
- Publisher.Window
ms.assetid: 342d77cd-5556-6ac3-a828-b1b60380f910
ms.date: 06/04/2019
localization_priority: Normal
---


# Window object (Publisher)

Represents a window. Many publication characteristics, such as scroll bars and rulers, are actually properties of the window.
 
## Remarks

Use the **[ActiveWindow](Publisher.Application.ActiveWindow.md)** property of the **Application** object to return a **Window** object.

Use the **Caption** property to return the file and application names of the active window. 

## Example

The following example maximizes the active window.

```vb
Sub MaximizeWindow 
 ActiveWindow.WindowState = pbWindowStateMaximize 
End Sub
```

<br/>

The following example displays a message with the file name and Microsoft Publisher application name.

```vb
Sub ShowFileApNames 
 MsgBox Windows(1).Caption 
End Sub
```

## Methods

- [Activate](Publisher.Window.Activate.md)
- [Move](Publisher.Window.Move.md)
- [Resize](Publisher.Window.Resize.md)

## Properties

- [Application](Publisher.Window.Application.md)
- [Caption](Publisher.Window.Caption.md)
- [Height](Publisher.Window.Height.md)
- [Hwnd](Publisher.Window.Hwnd.md)
- [Left](Publisher.Window.Left.md)
- [Parent](Publisher.Window.Parent.md)
- [Top](Publisher.Window.Top.md)
- [Visible](Publisher.Window.Visible.md)
- [Width](Publisher.Window.Width.md)
- [WindowState](Publisher.Window.WindowState.md)

## See also

- [Publisher Object Model Reference](overview/publisher/object-model.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]