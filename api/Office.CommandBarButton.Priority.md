---
title: CommandBarButton.Priority property (Office)
ms.prod: office
api_name:
- Office.CommandBarButton.Priority
ms.assetid: 72599580-16d2-20b3-05ad-b454afbba6ef
ms.date: 01/03/2019
localization_priority: Normal
---


# CommandBarButton.Priority property (Office)

Gets or sets the priority of a **CommandBarButton** control. A control's priority determines whether the control can be dropped from a docked command bar if the command bar controls can't fit in a single row. Controls that can't fit in a single row drop off command bars from right to left. Read/write.

> [!NOTE]
> The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, see [Overview of the Office Fluent ribbon](../library-reference/concepts/overview-of-the-office-fluent-ribbon.md).


## Syntax

_expression_.**Priority**

_expression_ A variable that represents a **[CommandBarButton](Office.CommandBarButton.md)** object.


## Remarks

Valid priority numbers are 0 (zero) through 7, and the default value is 3. A priority of 1 means that the control cannot be dropped from a toolbar. Other priority values are ignored.

The **Priority** property is not used by command bar controls that are menu items.


## See also

- [CommandBarButton object members](overview/library-reference/commandbarbutton-members-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]