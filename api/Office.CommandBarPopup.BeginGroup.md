---
title: CommandBarPopup.BeginGroup property (Office)
ms.prod: office
api_name:
- Office.CommandBarPopup.BeginGroup
ms.assetid: 0ecc5c98-5db7-792c-8f33-86f7df32d912
ms.date: 01/04/2019
localization_priority: Normal
---


# CommandBarPopup.BeginGroup property (Office)

Gets **True** if the specified command bar control appears at the beginning of a group of controls on the command bar. Read/write.

> [!NOTE] 
> The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, see [Overview of the Office Fluent ribbon](../library-reference/concepts/overview-of-the-office-fluent-ribbon.md).


## Syntax

_expression_.**BeginGroup**

_expression_ A variable that represents a **[CommandBarPopup](Office.CommandBarPopup.md)** object.


## Return value

Boolean


## Example

This example begins a new group with the last control on the active menu bar.


```vb
Set myMenuBar = CommandBars.ActiveMenuBar 
Set lastMenu = myMenuBar _ 
    .Controls(myMenuBar.Controls.Count) 
lastMenu.BeginGroup = True
```


## See also

- [CommandBarPopup object members](overview/library-reference/commandbarpopup-members-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]