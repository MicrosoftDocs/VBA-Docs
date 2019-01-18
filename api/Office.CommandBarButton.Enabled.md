---
title: CommandBarButton.Enabled property (Office)
ms.prod: office
api_name:
- Office.CommandBarButton.Enabled
ms.assetid: 264335ca-6506-0e86-16df-44af277ade83
ms.date: 01/03/2019
localization_priority: Normal
---


# CommandBarButton.Enabled property (Office)

**True** if the specified **CommandBar** or **CommandBarControl** is enabled. Read/write.

> [!NOTE]
> The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, see [Overview of the Office Fluent ribbon](../library-reference/concepts/overview-of-the-office-fluent-ribbon.md).


## Syntax

_expression_.**Enabled**

_expression_ A variable that represents a **[CommandBarButton](Office.CommandBarButton.md)** object.


## Return value

Boolean


## Remarks

For command bars, setting this property to **True** causes the name of the command bar to appear in the list of available command bars.

For built-in controls, if you set the **Enabled** property to **True**, the application determines its state, but setting it to **False** will force it to be disabled.


## See also

- [CommandBarButton object members](overview/library-reference/commandbarbutton-members-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]