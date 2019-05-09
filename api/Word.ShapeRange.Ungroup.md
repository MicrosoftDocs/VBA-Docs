---
title: ShapeRange.Ungroup method (Word)
keywords: vbawd10.chm162856987
f1_keywords:
- vbawd10.chm162856987
ms.prod: word
api_name:
- Word.ShapeRange.Ungroup
ms.assetid: 2a6b4eb1-724b-7ff8-5392-57dfdfaa815d
ms.date: 06/08/2017
localization_priority: Normal
---


# ShapeRange.Ungroup method (Word)

Ungroups any grouped shapes in the specified range of shapes, disassembles pictures and OLE objects within the specified shape or range of shapes, and returns the ungrouped shapes as a single  **ShapeRange** object.


## Syntax

_expression_.**Ungroup**

_expression_ Required. A variable that represents a **[ShapeRange](Word.shaperange.md)** object.


## Return value

ShapeRange


## Remarks

Because a group of shapes is treated as a single object, grouping and ungrouping shapes changes the number of items in the  **Shapes** collection and changes the index numbers of items that come after the ungrouped shapes in the collection.


## See also


[ShapeRange Collection Object](Word.shaperange.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]