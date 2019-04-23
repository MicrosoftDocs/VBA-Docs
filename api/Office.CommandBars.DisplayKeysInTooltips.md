---
title: CommandBars.DisplayKeysInTooltips property (Office)
keywords: vbaof11.chm2006
f1_keywords:
- vbaof11.chm2006
ms.prod: office
api_name:
- Office.CommandBars.DisplayKeysInTooltips
ms.assetid: de132c5f-bc9f-c335-28ff-b9459c912b2c
ms.date: 01/04/2019
localization_priority: Normal
---


# CommandBars.DisplayKeysInTooltips property (Office)

Is **True** if shortcut keys are displayed in the **ToolTips** for each command bar control. Read/write.

> [!NOTE] 
> The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, see [Overview of the Office Fluent ribbon](../library-reference/concepts/overview-of-the-office-fluent-ribbon.md).


## Syntax

_expression_.**DisplayKeysInTooltips**

_expression_ A variable that represents a **[CommandBars](Office.CommandBars.md)** object.


## Remarks

To display shortcut keys in **ToolTips**, you must also set the **DisplayTooltips** property to **True**.


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

- [CommandBars object members](overview/library-reference/commandbars-members-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]