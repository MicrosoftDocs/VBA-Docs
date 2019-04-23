---
title: CommandBarButton.PasteFace method (Office)
keywords: vbaof11.chm6004
f1_keywords:
- vbaof11.chm6004
ms.prod: office
api_name:
- Office.CommandBarButton.PasteFace
ms.assetid: 1c4179c4-b6b5-527f-5027-25ced8ee907d
ms.date: 01/03/2019
localization_priority: Normal
---


# CommandBarButton.PasteFace method (Office)

Pastes the contents of the Clipboard onto a **CommandBarButton**.

> [!NOTE]
> The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, see [Overview of the Office Fluent ribbon](../library-reference/concepts/overview-of-the-office-fluent-ribbon.md).


## Syntax

_expression_.**PasteFace**

_expression_ A variable that represents a **[CommandBarButton](Office.CommandBarButton.md)** object.


## Example

This example finds the built-in **FileOpen** button and pastes the face from the **Spelling and Grammar** button onto it from the Clipboard.


```vb
Set myControl = CommandBars.FindControl(Type:=msoControlButton, Id:=2) 
myControl.CopyFace 
Set myControl = CommandBars.FindControl(Type:=msoControlButton, Id:=23) 
myControl.PasteFace
```


## See also

- [CommandBarButton object members](overview/library-reference/commandbarbutton-members-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]