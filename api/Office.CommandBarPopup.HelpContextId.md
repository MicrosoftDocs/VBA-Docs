---
title: CommandBarPopup.HelpContextId property (Office)
ms.prod: office
api_name:
- Office.CommandBarPopup.HelpContextId
ms.assetid: b07d39b7-9fad-51dc-b093-de88cd1ea905
ms.date: 01/04/2019
localization_priority: Normal
---


# CommandBarPopup.HelpContextId property (Office)

Gets or sets the Help context Id number for the Help topic attached to the **CommandBarPopup** control. Read/write.

> [!NOTE] 
> The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, see [Overview of the Office Fluent ribbon](../library-reference/concepts/overview-of-the-office-fluent-ribbon.md).


## Syntax

_expression_.**HelpContextId**

_expression_ A variable that represents a **[CommandBarPopup](Office.CommandBarPopup.md)** object.


## Return value

Integer


## Remarks

To use this property, you must also set the **[HelpFile](office.commandbarpopup.helpfile.md)** property. Help topics respond to Shift+F1.


## See also

- [CommandBarPopup object members](overview/library-reference/commandbarpopup-members-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]