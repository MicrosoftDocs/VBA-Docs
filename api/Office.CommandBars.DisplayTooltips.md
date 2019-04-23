---
title: CommandBars.DisplayTooltips property (Office)
keywords: vbaof11.chm2005
f1_keywords:
- vbaof11.chm2005
ms.prod: office
api_name:
- Office.CommandBars.DisplayTooltips
ms.assetid: 98b62729-d1c8-a6dc-328e-8dbb6bbd80dc
ms.date: 01/04/2019
localization_priority: Normal
---


# CommandBars.DisplayTooltips property (Office)

Is **True** if **ScreenTips** are displayed whenever the user positions the pointer over command bar controls. Read/write.

> [!NOTE] 
>  The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, see [Overview of the Office Fluent ribbon](../library-reference/concepts/overview-of-the-office-fluent-ribbon.md).


## Syntax

_expression_.**DisplayTooltips**

_expression_ A variable that represents a **[CommandBars](Office.CommandBars.md)** object.


## Remarks

Setting the **DisplayTooltips** property in a container application immediately affects every command bar in every running Microsoft Office application, and in every Office application opened after the property is set.


## Example

This example displays large controls and **ToolTips** on all command bars.


```vb
Set allBars = CommandBars 
 
allBars.LargeButtons = True  
allBars.DisplayTooltips = True  

```


## See also

- [CommandBars object members](overview/library-reference/commandbars-members-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]