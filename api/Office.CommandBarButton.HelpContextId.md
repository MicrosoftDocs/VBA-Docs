---
title: CommandBarButton.HelpContextId property (Office)
ms.prod: office
api_name:
- Office.CommandBarButton.HelpContextId
ms.assetid: 2e4f33db-7143-dd8d-65b3-d0c993f2e966
ms.date: 01/03/2019
localization_priority: Normal
---


# CommandBarButton.HelpContextId property (Office)

Gets or sets the Help context Id number for the Help topic attached to the **CommandBarButton** control. Read/write.

> [!NOTE]
> The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, see [Overview of the Office Fluent ribbon](../library-reference/concepts/overview-of-the-office-fluent-ribbon.md).


## Syntax

_expression_.**HelpContextId**

_expression_ A variable that represents a **[CommandBarButton](Office.CommandBarButton.md)** object.


## Return value

Integer


## Remarks

To use this property, you must also set the **[HelpFile](office.commandbarbutton.helpfile.md)** property. Help topics respond to Shift+F1.


## See also

- [CommandBarButton object members](overview/library-reference/commandbarbutton-members-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]