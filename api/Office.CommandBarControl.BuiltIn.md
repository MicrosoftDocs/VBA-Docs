---
title: CommandBarControl.BuiltIn property (Office)
ms.prod: office
api_name:
- Office.CommandBarControl.BuiltIn
ms.assetid: 4b3904dc-3376-28e0-6c93-4acff8101e6f
ms.date: 01/04/2019
localization_priority: Normal
---


# CommandBarControl.BuiltIn property (Office)

Gets **True** if the specified command bar control is a built-in control of the container application. Returns **False** if it's a custom control, or if it's a built-in control whose **OnAction** property has been set. Read-only.

> [!NOTE] 
> The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, see [Overview of the Office Fluent ribbon](../library-reference/concepts/overview-of-the-office-fluent-ribbon.md).


## Syntax

_expression_.**BuiltIn**

_expression_ A variable that represents a **[CommandBarControl](Office.CommandBarControl.md)** object.


## Return value

Boolean


## See also

- [CommandBarControl object members](overview/library-reference/commandbarcontrol-members-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]