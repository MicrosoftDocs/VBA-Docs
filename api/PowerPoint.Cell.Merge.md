---
title: Cell.Merge method (PowerPoint)
keywords: vbapp10.chm628005
f1_keywords:
- vbapp10.chm628005
ms.prod: powerpoint
api_name:
- PowerPoint.Cell.Merge
ms.assetid: e4830df1-4db9-f1e0-a4c6-d4ed2d99b9fa
ms.date: 06/08/2017
localization_priority: Normal
---


# Cell.Merge method (PowerPoint)

Merges one table cell with another. The result is a single table cell.


## Syntax

_expression_.**Merge** (_MergeTo_)

_expression_ A variable that represents a [Cell](PowerPoint.Cell.md) object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _MergeTo_|Required|**Cell**|The **Cell** object to be merged with.|

## Remarks

For the MergeTo parameter, use the syntax  `.Cell(row, column)`.

This method returns an error if the file name cannot be opened, or the presentation has a baseline.


## Example

This example merges the first two cells of row one in the specified table.


```vb
With ActivePresentation.Slides(2).Shapes(5).Table

    .Cell(1, 1).Merge MergeTo:=.Cell(1, 2)

End With
```


## See also


[Presentation Object](PowerPoint.Presentation.md)
[Cell Object](PowerPoint.Cell.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]