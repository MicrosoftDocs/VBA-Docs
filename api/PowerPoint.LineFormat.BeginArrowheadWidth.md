---
title: LineFormat.BeginArrowheadWidth property (PowerPoint)
keywords: vbapp10.chm553005
f1_keywords:
- vbapp10.chm553005
ms.prod: powerpoint
api_name:
- PowerPoint.LineFormat.BeginArrowheadWidth
ms.assetid: 3834e2c8-d153-57f8-014e-1545326dd370
ms.date: 06/08/2017
localization_priority: Normal
---


# LineFormat.BeginArrowheadWidth property (PowerPoint)

Returns or sets the width of the arrowhead at the beginning of the specified line. Read/write.


## Syntax

_expression_.**BeginArrowheadWidth**

_expression_ A variable that represents a [LineFormat](PowerPoint.LineFormat.md) object.


## Return value

MsoArrowheadWidth


## Remarks

The value of the  **BeginArrowheadWidth** property can be one of these **MsoArrowheadWidth** constants


||
|:-----|
|**msoArrowheadNarrow**|
|**msoArrowheadWide**|
|**msoArrowheadWidthMedium**|
|**msoArrowheadWidthMixed**|

## Example

This example adds a line to _myDocument_. There's a short, narrow oval on the line's starting point and a long, wide triangle on its endpoint.


```vb
Set myDocument = ActivePresentation.Slides(1)

With myDocument.Shapes.AddLine(100, 100, 200, 300).Line

    .BeginArrowheadLength = msoArrowheadShort

    .BeginArrowheadStyle = msoArrowheadOval

    .BeginArrowheadWidth = msoArrowheadNarrow

    .EndArrowheadLength = msoArrowheadLong

    .EndArrowheadStyle = msoArrowheadTriangle

    .EndArrowheadWidth = msoArrowheadWide

End With
```


## See also


[LineFormat Object](PowerPoint.LineFormat.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]