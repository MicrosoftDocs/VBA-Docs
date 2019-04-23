---
title: CommandBarButton.CopyFace method (Office)
keywords: vbaof11.chm6002
f1_keywords:
- vbaof11.chm6002
ms.prod: office
api_name:
- Office.CommandBarButton.CopyFace
ms.assetid: 09f09dbd-b70f-8b7d-1af7-7e43bffe3030
ms.date: 01/03/2019
localization_priority: Normal
---


# CommandBarButton.CopyFace method (Office)

Copies the face of a command bar button control to the Clipboard.

> [!NOTE]
> The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, see [Overview of the Office Fluent ribbon](../library-reference/concepts/overview-of-the-office-fluent-ribbon.md).


## Syntax

_expression_.**CopyFace**

_expression_ A variable that represents a **[CommandBarButton](Office.CommandBarButton.md)** object.


## Remarks

Use the **PasteFace** method to paste the contents of the Clipboard onto a button face.


## Example

This example finds the built-in **Open** button, copies the button face to the Clipboard, and then pastes the face onto the **Spelling** and **Grammar** button.


```vb
Set myControl = CommandBars.FindControl(Type:=msoControlButton, Id:=23) 
myControl.CopyFace 
Set myControl = CommandBars.FindControl(Type:=msoControlButton, ID:=2) 
myControl.PasteFace
```


## See also

- [CommandBarButton object members](overview/library-reference/commandbarbutton-members-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]