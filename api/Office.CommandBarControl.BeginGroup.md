---
title: CommandBarControl.BeginGroup property (Office)
ms.prod: office
api_name:
- Office.CommandBarControl.BeginGroup
ms.assetid: 529b8c23-ec1f-b37b-a40c-9ae6016f4dc0
ms.date: 01/04/2019
localization_priority: Normal
---


# CommandBarControl.BeginGroup property (Office)

Gets **True** if the specified command bar control appears at the beginning of a group of controls on the command bar. Read/write.

> [!NOTE] 
> The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, see [Overview of the Office Fluent ribbon](../library-reference/concepts/overview-of-the-office-fluent-ribbon.md).


## Syntax

_expression_.**BeginGroup**

_expression_ A variable that represents a **[CommandBarControl](Office.CommandBarControl.md)** object.


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

- [CommandBarControl object members](overview/library-reference/commandbarcontrol-members-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]