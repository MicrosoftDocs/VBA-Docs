---
title: CommandBarButton.ShortcutText property (Office)
keywords: vbaof11.chm6005
f1_keywords:
- vbaof11.chm6005
ms.prod: office
api_name:
- Office.CommandBarButton.ShortcutText
ms.assetid: e0c76e70-16db-d3ae-9767-069579c8ea91
ms.date: 01/03/2019
localization_priority: Normal
---


# CommandBarButton.ShortcutText property (Office)

Gets or sets the shortcut key text displayed next to a **CommandBarButton** control when the button appears on a menu, submenu, or shortcut menu. Read/write.

> [!NOTE]
> The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, see [Overview of the Office Fluent ribbon](../library-reference/concepts/overview-of-the-office-fluent-ribbon.md).


## Syntax

_expression_.**ShortcutText**

_expression_ A variable that represents a **[CommandBarButton](Office.CommandBarButton.md)** object.


## Remarks

You can set this property only for command bar buttons that contain an **OnAction** macro.


## Example

This example displays the shortcut text for the **Open** command (**File** menu) on the Microsoft Excel Worksheet menu bar in a message box.


```vb
MsgBox (CommandBars("Worksheet Menu Bar"). _ 
    Controls("File").Controls("New...).ShortcutText)
```


## See also

- [CommandBarButton object members](overview/library-reference/commandbarbutton-members-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]