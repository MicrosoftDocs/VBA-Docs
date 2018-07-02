---
title: CommandBar.Left Property (Office)
keywords: vbaof11.chm3009
f1_keywords:
- vbaof11.chm3009
ms.prod: office
api_name:
- Office.CommandBar.Left
ms.assetid: 2353aef6-aaa1-76b9-33da-57bbe1df30af
ms.date: 06/08/2017
---


# CommandBar.Left Property (Office)

Sets or gets the horizontal distance (in pixels) of the  **CommandBar** from the left edge of the object relative to the screen. Read/write.

> [!NOTE]
> The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, search Help for the keyword "ribbon."


## Syntax

 _expression_. `Left`

 _expression_ Required. A variable that represents a '[CommandBar](Office.CommandBar.md)' object.


## Example

This example moves the command bar named Custom from its docked position along the top of the window to the left edge of the window.


```vb
Set myBar = CommandBars("Custom") 
With myBar 
    .Position = 1 
    .RowIndex = 2 
    .Left = 0 
End With
```


## See also


[CommandBar Object](Office.CommandBar.md)



[CommandBar Object Members](commandbar-members-office.md)

