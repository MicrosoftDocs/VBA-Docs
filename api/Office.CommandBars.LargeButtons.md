---
title: CommandBars.LargeButtons property (Office)
keywords: vbaof11.chm2009
f1_keywords:
- vbaof11.chm2009
ms.prod: office
api_name:
- Office.CommandBars.LargeButtons
ms.assetid: bcacab92-9779-5061-f68a-69722210e14e
ms.date: 01/04/2019
localization_priority: Normal
---


# CommandBars.LargeButtons property (Office)

Is **True** if the toolbar buttons displayed are larger than normal size. Read/write.

> [!NOTE] 
> The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, see [Overview of the Office Fluent ribbon](../library-reference/concepts/overview-of-the-office-fluent-ribbon.md).


## Syntax

_expression_.**LargeButtons**

_expression_ A variable that represents a **[CommandBars](Office.CommandBars.md)** object.


## Example

This example switches the display size of toolbar buttons on all command bars.


```vb
Set allBars = CommandBars 
If allBars.LargeButtons Then 
    allBars.LargeButtons = False  
Else 
    allBars.LargeButtons = True  
End If
```


## See also

- [CommandBars object members](overview/library-reference/commandbars-members-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]