---
title: CommandBarPopup.Index property (Office)
ms.prod: office
api_name:
- Office.CommandBarPopup.Index
ms.assetid: 6f6f6d1f-a59a-cf52-d273-a732652b4f05
ms.date: 01/04/2019
localization_priority: Normal
---


# CommandBarPopup.Index property (Office)

Gets a **Long** representing the index number for a **CommandBarPopup** object in the collection. Read-only.

> [!NOTE] 
> The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, see [Overview of the Office Fluent ribbon](../library-reference/concepts/overview-of-the-office-fluent-ribbon.md).


## Syntax

_expression_.**Index**

_expression_ A variable that represents a **[CommandBarPopup](Office.CommandBarPopup.md)** object.


## Return value

Long


## Remarks

The position of the first command bar control is 1. Separators are not counted in the **CommandBarControls** collection.


## See also

- [CommandBarPopup object members](overview/library-reference/commandbarpopup-members-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]