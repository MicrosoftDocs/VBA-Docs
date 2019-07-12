---
title: LineFormat.BeginArrowheadLength property (PowerPoint)
keywords: vbapp10.chm553003
f1_keywords:
- vbapp10.chm553003
ms.prod: powerpoint
api_name:
- PowerPoint.LineFormat.BeginArrowheadLength
ms.assetid: b46151e1-251f-7498-9dfc-b652b356edf0
ms.date: 06/08/2017
localization_priority: Normal
---


# LineFormat.BeginArrowheadLength property (PowerPoint)

Returns or sets the length of the arrowhead at the beginning of the specified line. Read/write.


## Syntax

_expression_.**BeginArrowheadLength**

_expression_ A variable that represents a [LineFormat](PowerPoint.LineFormat.md) object.


## Return value

MsoArrowheadLength


## Remarks

The value of the  **BeginArrowheadLength** property can be one of these **MsoArrowheadLength** constants


||
|:-----|
|**msoArrowheadLengthMedium**|
|**msoArrowheadLengthMixed**|
|**msoArrowheadLong**|
|**msoArrowheadShort**|

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