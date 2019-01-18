---
title: CommandBar.Top property (Office)
keywords: vbaof11.chm3018
f1_keywords:
- vbaof11.chm3018
ms.prod: office
api_name:
- Office.CommandBar.Top
ms.assetid: 1bac668a-0caa-d185-cc07-ba55809c79fe
ms.date: 01/03/2019
localization_priority: Normal
---


# CommandBar.Top property (Office)

Sets or gets the distance from the top edge of the specified command bar, to the top edge of the screen. For docked command bars, this property returns or sets the distance from the command bar to the top of the docking area. Read/write.

> [!NOTE]
> The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, see [Overview of the Office Fluent ribbon](../library-reference/concepts/overview-of-the-office-fluent-ribbon.md).


## Syntax

_expression_.**Top**

_expression_ Required. A variable that represents a **[CommandBar](Office.CommandBar.md)** object.


## Example

This example positions the upper-left corner of the floating command bar named **Custom** 140 pixels from the left edge of the screen and 100 pixels from the top of the screen.


```vb
Set myBar = CommandBars("Custom") 
myBar.Position = msoBarFloating 
With myBar 
    .Left = 140 
    .Top = 100 
End With
```


## See also

- [CommandBar object members](overview/library-reference/commandbar-members-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]