---
title: CommandBars.MenuAnimationStyle property (Office)
keywords: vbaof11.chm2010
f1_keywords:
- vbaof11.chm2010
ms.prod: office
api_name:
- Office.CommandBars.MenuAnimationStyle
ms.assetid: bd79a55a-23f4-6056-649b-9dc384b597aa
ms.date: 06/08/2017
---


# CommandBars.MenuAnimationStyle property (Office)

Gets or sets a  **MsoMenuAnimation** that represents the way a command bar is animated. Read/write.

> [!NOTE] 
> The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, search Help for the keyword "ribbon."


## Syntax

_expression_. `MenuAnimationStyle`

_expression_ A variable that represents a [CommandBars](Office.CommandBars.md) object.


## Example

This example sets options for all command bars in Microsoft Office.


```vb
With CommandBars 
    .LargeButtons = True  
    .DisplayTooltips = True  
    .DisplayKeysInTooltips = True  
    .MenuAnimationStyle = msoMenuAnimationUnfold 
End With
```


## See also


[CommandBars Object](Office.CommandBars.md)



[CommandBars Object Members](./overview/Library-Reference/commandbars-members-office.md)

