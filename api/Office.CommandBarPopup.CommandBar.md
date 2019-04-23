---
title: CommandBarPopup.CommandBar property (Office)
keywords: vbaof11.chm7001
f1_keywords:
- vbaof11.chm7001
ms.prod: office
api_name:
- Office.CommandBarPopup.CommandBar
ms.assetid: e78abe18-d260-8cac-d647-322b449e4bbb
ms.date: 01/04/2019
localization_priority: Normal
---


# CommandBarPopup.CommandBar property (Office)

Gets a **CommandBar** object that represents the menu displayed by the specified popup control. Read-only.

> [!NOTE] 
> The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, see [Overview of the Office Fluent ribbon](../library-reference/concepts/overview-of-the-office-fluent-ribbon.md).


## Syntax

_expression_.**CommandBar**

_expression_ A variable that represents a **[CommandBarPopup](Office.CommandBarPopup.md)** object.


## Example

This example sets the variable `fourthLevel` to the fourth control on the command bar named **Drawing**.


```vb
Set fourthLevel = CommandBars("Drawing") _ 
    .Controls(1).CommandBar.Controls(4)
```


## See also

- [CommandBarPopup object members](overview/library-reference/commandbarpopup-members-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]