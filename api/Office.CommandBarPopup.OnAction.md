---
title: CommandBarPopup.OnAction property (Office)
ms.prod: office
api_name:
- Office.CommandBarPopup.OnAction
ms.assetid: 47511647-5f1f-5e40-179b-ec589a2c39be
ms.date: 01/04/2019
localization_priority: Normal
---


# CommandBarPopup.OnAction property (Office)

Gets or sets the name of a Visual Basic procedure that will run when the user clicks or changes the value of a **CommandBarPopup** control. Read/write.

> [!NOTE] 
> The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, see [Overview of the Office Fluent ribbon](../library-reference/concepts/overview-of-the-office-fluent-ribbon.md).


## Syntax

_expression_.**OnAction**

_expression_ A variable that represents a **[CommandBarPopup](Office.CommandBarPopup.md)** object.


## Return value

String


## Remarks

The container application determines whether the value is a valid macro name.


## See also

- [CommandBarPopup object members](overview/library-reference/commandbarpopup-members-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]