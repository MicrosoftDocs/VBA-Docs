---
title: Row.Height property (PowerPoint)
keywords: vbapp10.chm626006
f1_keywords:
- vbapp10.chm626006
ms.prod: powerpoint
api_name:
- PowerPoint.Row.Height
ms.assetid: a4334eed-66c3-0042-585d-069ce23ffb3d
ms.date: 06/08/2017
localization_priority: Normal
---


# Row.Height property (PowerPoint)

Returns or sets the height of the specified object, in points. Read/write.


## Syntax

_expression_.**Height**

_expression_ A variable that represents a [Row](PowerPoint.Row.md) object.


## Return value

Single


## Remarks

The **Height** property of a **Shape** object returns or sets the height of the forward-facing surface of the specified shape. This measurement doesn't include shadows or 3D effects.


## Example

This example sets the height of document window two to half the height of the application window.


```vb
Windows(2).Height = Application.Height / 2
```

This example sets the height for row two in the specified table to 100 points (72 points per inch).




```vb
ActivePresentation.Slides(2).Shapes(5).Table.Rows(2).Height = 100
```


## See also


[Row Object](PowerPoint.Row.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]