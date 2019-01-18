---
title: CommandBar.RowIndex property (Office)
keywords: vbaof11.chm3014
f1_keywords:
- vbaof11.chm3014
ms.prod: office
api_name:
- Office.CommandBar.RowIndex
ms.assetid: 6dd5576c-0a46-9a72-9c4e-fcf685097b77
ms.date: 01/03/2019
localization_priority: Normal
---


# CommandBar.RowIndex property (Office)

Gets or sets the docking order of a command bar in relation to other command bars in the same docking area. Can be an integer greater than zero, or either of the following **[msoBarRow](office.msobarrow.md)** constants: **msoBarRowFirst** or **msoBarRowLast**. Read/write.

> [!NOTE]
> The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, see [Overview of the Office Fluent ribbon](../library-reference/concepts/overview-of-the-office-fluent-ribbon.md).


## Syntax

_expression_.**RowIndex**

_expression_ A variable that represents a **[CommandBar](Office.CommandBar.md)** object.


## Remarks

Several command bars can share the same row index, and command bars with lower numbers are docked first. If two or more command bars share the same row index, the command bar most recently assigned will be displayed first in its group.


## Example

This example adjusts the position of the command bar named **Custom** by moving it to the left 110 pixels more than the default, and it makes this command bar the first to be docked by changing its row index to **msoBarRowFirst**.


```vb
Set myBar = CommandBars("Custom") 
With myBar 
    .RowIndex = msoBarRowFirst 
    .Left = 140 
End With
```


## See also

- [CommandBar object members](overview/library-reference/commandbar-members-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]