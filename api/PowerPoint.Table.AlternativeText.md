---
title: Table.AlternativeText property (PowerPoint)
keywords: vbapp10.chm622018
f1_keywords:
- vbapp10.chm622018
ms.prod: powerpoint
api_name:
- PowerPoint.Table.AlternativeText
ms.assetid: db35ce8c-0115-4e72-db25-3d555242aee4
ms.date: 06/08/2017
localization_priority: Normal
---


# Table.AlternativeText property (PowerPoint)

Returns or sets the alternative text associated with a shape in a Web presentation. Read/write.


## Syntax

_expression_.**AlternativeText**

_expression_ A variable that represents a [Table](PowerPoint.Table.md) object.


## Example

The following example sets the alternative text for the selected shape in the active window. The selected shape is a picture of a mallard duck.


```vb
ActiveWindow.Selection.ShapeRange.AlternativeText = "This is a mallard duck."
```


## See also


[Table Object](PowerPoint.Table.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]