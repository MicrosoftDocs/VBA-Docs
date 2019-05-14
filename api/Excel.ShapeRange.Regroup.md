---
title: ShapeRange.Regroup method (Excel)
keywords: vbaxl10.chm640089
f1_keywords:
- vbaxl10.chm640089
ms.prod: excel
api_name:
- Excel.ShapeRange.Regroup
ms.assetid: d30d3064-c37e-84b0-10a6-11dcd18c593e
ms.date: 05/14/2019
localization_priority: Normal
---


# ShapeRange.Regroup method (Excel)

Regroups the group that the specified shape range belonged to previously. Returns the regrouped shapes as a single **[Shape](Excel.Shape.md)** object.


## Syntax

_expression_.**Regroup**

_expression_ A variable that represents a **[ShapeRange](Excel.shaperange.md)** object.


## Return value

Shape


## Remarks

The **Regroup** method only restores the group for the first previously grouped shape that it finds in the specified **ShapeRange** collection. Therefore, if the specified shape range contains shapes that previously belonged to different groups, only one of the groups will be restored.

Note that because a group of shapes is treated as a single shape, grouping and ungrouping shapes changes the number of items in the **[Shapes](Excel.Shapes.md)** collection and changes the index numbers of items that come after the affected items in the collection.


## Example

This example regroups the shapes in the selection in the active window. If the shapes haven't been previously grouped and ungrouped, this example will fail.

```vb
ActiveWindow.Selection.ShapeRange.Regroup
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]