---
title: Chart.AlternativeText property (PowerPoint)
keywords: vbapp10.chm684054
f1_keywords:
- vbapp10.chm684054
ms.prod: powerpoint
api_name:
- PowerPoint.Chart.AlternativeText
ms.assetid: bdded8b9-5f6e-dd83-db04-0ce180bd2552
ms.date: 06/08/2017
localization_priority: Normal
---


# Chart.AlternativeText property (PowerPoint)

Returns or sets the alternative text associated with a shape in a Web presentation. Read/write.


## Syntax

_expression_.**AlternativeText**

_expression_ A variable that represents a [Chart](PowerPoint.Chart.md) object.


## Return value

 **String**


## Example

The following example sets the alternative text for the selected shape in the active window. The selected shape is a picture of a mallard duck.


```vb
ActiveWindow.Selection.ShapeRange.AlternativeText = "This is a mallard duck."
```


## See also


[Chart Object](PowerPoint.Chart.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]