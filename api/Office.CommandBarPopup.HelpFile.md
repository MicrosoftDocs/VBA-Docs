---
title: CommandBarPopup.HelpFile property (Office)
ms.prod: office
api_name:
- Office.CommandBarPopup.HelpFile
ms.assetid: 67c79cb5-cca7-d113-49de-9f636c757867
ms.date: 01/04/2019
localization_priority: Normal
---


# CommandBarPopup.HelpFile property (Office)

Gets or sets the file name for the Help topic attached to the **CommandBarPopup** control. Read/write.

> [!NOTE] 
> The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, see [Overview of the Office Fluent ribbon](../library-reference/concepts/overview-of-the-office-fluent-ribbon.md).


## Syntax

_expression_.**HelpFile**

_expression_ A variable that represents a **[CommandBarPopup](Office.CommandBarPopup.md)** object.


## Return value

String


## Remarks

To use this property, you must also set the **[HelpContextID](office.commandbarpopup.helpcontextid.md)** property. Help topics respond to the user pressing Shift+F1.


## See also

- [CommandBarPopup object members](overview/library-reference/commandbarpopup-members-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]