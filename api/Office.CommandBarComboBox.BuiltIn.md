---
title: CommandBarComboBox.BuiltIn property (Office)
ms.prod: office
api_name:
- Office.CommandBarComboBox.BuiltIn
ms.assetid: 4dc0232c-94dd-ce40-95cd-7700fdd9a427
ms.date: 01/03/2019
localization_priority: Normal
---


# CommandBarComboBox.BuiltIn property (Office)

Gets **True** if the specified command bar control is a built-in control of the container application. Returns **False** if it's a custom control, or if it's a built-in control whose **OnAction** property has been set. Read-only.

> [!NOTE]
> The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, see [Overview of the Office Fluent ribbon](../library-reference/concepts/overview-of-the-office-fluent-ribbon.md).


## Syntax

_expression_.**BuiltIn**

_expression_ A variable that represents a **[CommandBarComboBox](Office.CommandBarComboBox.md)** object.


## Return value

Boolean


## See also

- [CommandBarComboBox object members](overview/library-reference/commandbarcombobox-members-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]