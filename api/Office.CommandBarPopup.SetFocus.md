---
title: CommandBarPopup.SetFocus method (Office)
ms.prod: office
api_name:
- Office.CommandBarPopup.SetFocus
ms.assetid: ce132a0d-aa1f-c8b1-2697-1cfe78b99123
ms.date: 06/08/2017
---


# CommandBarPopup.SetFocus method (Office)

Moves the keyboard focus to the specified  **CommandBarPopup** control. If the popup is disabled or isn't visible, this method will fail.

> [!NOTE] 
> The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, search Help for the keyword "ribbon."


## Syntax

_expression_. `SetFocus`

_expression_ A variable that represents a [CommandBarPopup](Office.CommandBarPopup.md) object.


## Example

The following example sets a reference to an existing command bar popup and then resets it to its default state.


```vb
Dim cbPopup As CommandBarPopup 
Set cbPopup = Application.CommandBars.FindControl _ 
(Type:=msoControlPopup, Tag:="Graphics") 
cbPopup.Reset 

```


## See also


[CommandBarPopup Object](Office.CommandBarPopup.md)



[CommandBarPopup Object Members](./overview/Library-Reference/commandbarpopup-members-office.md)

