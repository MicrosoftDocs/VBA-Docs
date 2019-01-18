---
title: CommandBarButton.HelpFile property (Office)
ms.prod: office
api_name:
- Office.CommandBarButton.HelpFile
ms.assetid: 6e97a52d-f50d-600b-26eb-b22988bd5ed5
ms.date: 01/03/2019
localization_priority: Normal
---


# CommandBarButton.HelpFile property (Office)

Gets or sets the file name for the Help topic attached to the **CommandBarButton** control. Read/write.

> [!NOTE]
> The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, see [Overview of the Office Fluent ribbon](../library-reference/concepts/overview-of-the-office-fluent-ribbon.md).


## Syntax

_expression_.**HelpFile**

_expression_ A variable that represents a **[CommandBarButton](Office.CommandBarButton.md)** object.


## Return value

String


## Remarks

To use this property, you must also set the **[HelpContextID](office.commandbarbutton.helpcontextid.md)** property. Help topics respond to the user pressing SHIFT+F1.


## See also

- [CommandBarButton object members](overview/library-reference/commandbarbutton-members-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]