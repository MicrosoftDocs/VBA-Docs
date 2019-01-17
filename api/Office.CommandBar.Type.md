---
title: CommandBar.Type property (Office)
keywords: vbaof11.chm3019
f1_keywords:
- vbaof11.chm3019
ms.prod: office
api_name:
- Office.CommandBar.Type
ms.assetid: e023edd9-a8f4-c20f-c6b1-c434182bd748
ms.date: 01/03/2019
localization_priority: Normal
---


# CommandBar.Type property (Office)

Gets the type of command bar. Read-only.

> [!NOTE]
> The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, see [Overview of the Office Fluent ribbon](../library-reference/concepts/overview-of-the-office-fluent-ribbon.md).


## Syntax

_expression_.**Type**

_expression_ Required. A variable that represents a **[CommandBar](Office.CommandBar.md)** object.


## Example

This example finds the first control on the command bar named **Custom**. Using the **Type** property, the example determines whether the control is a button. If the control is a button, the example copies the face of the **Copy** button (on the **Standard** toolbar), and then pastes it onto the control.

```vb
Set oldCtrl = CommandBars("Custom").Controls(1) 
If oldCtrl.Type = msoControlButton Then 
    Set newCtrl = CommandBars.FindControl(Type:= _ 
        MsoControlButton, ID:= _ 
        CommandBars("Standard").Controls("Copy").ID) 
    NewCtrl.CopyFace 
    OldCtrl.PasteFace 
End If
```


## See also

- [CommandBar object members](overview/library-reference/commandbar-members-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]