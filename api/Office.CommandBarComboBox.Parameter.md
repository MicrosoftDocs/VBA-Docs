---
title: CommandBarComboBox.Parameter property (Office)
ms.prod: office
api_name:
- Office.CommandBarComboBox.Parameter
ms.assetid: b5019fba-5124-5d9c-7abe-db10df32078b
ms.date: 01/03/2019
localization_priority: Normal
---


# CommandBarComboBox.Parameter property (Office)

Gets or sets a string that an application can use to execute a command from a **CommandBarComboBox** control. Read/write.

> [!NOTE]
> The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, see [Overview of the Office Fluent ribbon](../library-reference/concepts/overview-of-the-office-fluent-ribbon.md).


## Syntax

_expression_.**Parameter**

_expression_ A variable that represents a **[CommandBarComboBox](Office.CommandBarComboBox.md)** object.


## Return value

String


## Remarks

If the specified parameter is set for a built-in control, the application can modify its default behavior if it can parse and use the new value. If the parameter is set for custom controls, it can be used to send information to Visual Basic procedures, or it can be used to hold information about the control (similar to a second **Tag** property value).


## See also

- [CommandBarComboBox object members](overview/library-reference/commandbarcombobox-members-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]