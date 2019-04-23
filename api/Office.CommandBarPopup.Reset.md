---
title: CommandBarPopup.Reset method (Office)
ms.prod: office
api_name:
- Office.CommandBarPopup.Reset
ms.assetid: 8e31b4e2-66d1-b902-f837-dc4833b1607f
ms.date: 01/04/2019
localization_priority: Normal
---


# CommandBarPopup.Reset method (Office)

Resets a built-in **CommandBarPopup** control to its original function and face.

> [!NOTE] 
> The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, see [Overview of the Office Fluent ribbon](../library-reference/concepts/overview-of-the-office-fluent-ribbon.md).


## Syntax

_expression_.**Reset**

_expression_ A variable that represents a **[CommandBarPopup](Office.CommandBarPopup.md)** object.


## Remarks

Resetting a built-in control restores the actions originally intended for the control and resets each of the control's properties back to its original state. 


## Example

The following example searches all command bars for a **CommandBarPopup** object whose tag is **Graphics** and then resets it to its default state.


```vb
Set myControl = Application.CommandBars.FindControl _ 
(Type:=msoControlPopup, Tag:="Graphics")  
myControl.Reset 

```


## See also

- [CommandBarPopup object members](overview/library-reference/commandbarpopup-members-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]