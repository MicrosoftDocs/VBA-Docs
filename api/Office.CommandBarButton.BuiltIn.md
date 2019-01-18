---
title: CommandBarButton.BuiltIn property (Office)
ms.prod: office
api_name:
- Office.CommandBarButton.BuiltIn
ms.assetid: 0a159c65-99d1-efdf-ec5c-f4e51060dd09
ms.date: 01/03/2019
localization_priority: Normal
---


# CommandBarButton.BuiltIn property (Office)

Is **True** if the specified command bar control is a control of the container application. Returns **False** if it's a custom control, or if it's a built-in control whose **OnAction** property has been set. Read-only.

> [!NOTE]
> The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, see [Overview of the Office Fluent ribbon](../library-reference/concepts/overview-of-the-office-fluent-ribbon.md).


## Syntax

_expression_.**BuiltIn**

_expression_ A variable that represents a **[CommandBarButton](Office.CommandBarButton.md)** object.


## Return value

Boolean


## See also

- [CommandBarButton object members](overview/library-reference/commandbarbutton-members-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]