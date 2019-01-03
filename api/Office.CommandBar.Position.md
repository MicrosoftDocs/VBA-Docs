---
title: CommandBar.Position property (Office)
keywords: vbaof11.chm3013
f1_keywords:
- vbaof11.chm3013
ms.prod: office
api_name:
- Office.CommandBar.Position
ms.assetid: b1e80bc0-1586-523b-a9ec-70c76fa54252
ms.date: 06/08/2017
---


# CommandBar.Position property (Office)

Gets or sets a  **MsoBarPosition** constant representing the position of a command bar. Read/write.

> [!NOTE]
> The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, search Help for the keyword "ribbon."


## Syntax

 _expression_. `Position`

 _expression_ A variable that represents a [CommandBar](Office.CommandBar.md) object.


## Example

This example steps through the collection of command bars, docking the custom command bars at the bottom of the application window and docking the built-in command bars at the top of the window.


```vb
For Each bar In CommandBars 
    If bar.Visible = True Then 
        If bar.BuiltIn Then 
            bar.Position = msoBarTop 
         Else 
            bar.Position = msoBarBottom 
        End If 
    End If 
Next
```


## See also


[CommandBar Object](Office.CommandBar.md)



[CommandBar Object Members](./overview/Library-Reference/commandbar-members-office.md)

