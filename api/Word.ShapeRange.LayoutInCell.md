---
title: ShapeRange.LayoutInCell property (Word)
keywords: vbawd10.chm162857105
f1_keywords:
- vbawd10.chm162857105
ms.prod: word
api_name:
- Word.ShapeRange.LayoutInCell
ms.assetid: ed09bd81-007c-0907-5eea-e9e3edf70e8b
ms.date: 06/08/2017
localization_priority: Normal
---


# ShapeRange.LayoutInCell property (Word)

Returns a  **Long** that represents whether a shape in a table is displayed inside the table or outside the table. .


## Syntax

_expression_. `LayoutInCell`

_expression_ Required. A variable that represents a **[ShapeRange](Word.shaperange.md)** object.


## Remarks

The  **LayoutInCell** property corresponds to the **Layout in table cell** option in the **Advanced Layout** dialog box for picture formatting. **True** indicates that a specified picture is displayed within the table. **False** indicates that a specified picture is displayed outside the table.


> [!NOTE] 
> Setting the  **LayoutInCell** property will take effect only if the **Type** property of the **WrapFormat** object is set to something other than **wdWrapTypeInline** or **wdWrapTypeNone**.


## See also


[ShapeRange Collection Object](Word.shaperange.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]