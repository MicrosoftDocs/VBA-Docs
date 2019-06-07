---
title: LineFormat.EndArrowheadStyle property (Publisher)
keywords: vbapb10.chm3408134
f1_keywords:
- vbapb10.chm3408134
ms.prod: publisher
api_name:
- Publisher.LineFormat.EndArrowheadStyle
ms.assetid: 991354c7-3f2c-a882-74d6-1c5cd3019494
ms.date: 06/08/2019
localization_priority: Normal
---


# LineFormat.EndArrowheadStyle property (Publisher)

Returns or sets an **[MsoArrowheadStyle](Office.MsoArrowheadStyle.md)** constant indicating the style of the arrowhead at the end of the specified line. Read/write.


## Syntax

_expression_.**EndArrowheadStyle**

_expression_ A variable that represents a **[LineFormat](Publisher.LineFormat.md)** object.


## Return value

MsoArrowheadStyle


## Remarks

Use the **[BeginArrowheadStyle](Publisher.LineFormat.BeginArrowheadStyle.md)** property to return or set the style of the arrowhead at the beginning of the line.

The **EndArrowheadStyle** property value can be one of the **MsoArrowheadStyle** constants declared in the Microsoft Office type library.


## Example

This example adds a line to the active publication. There is a short, narrow oval on the line's starting point and a long, wide triangle on its endpoint.

```vb
With ActiveDocument.Pages(1).Shapes _ 
 .AddLine(BeginX:=100, BeginY:=100, _ 
 EndX:=200, EndY:=300).Line 
 .BeginArrowheadLength = msoArrowheadShort 
 .BeginArrowheadStyle = msoArrowheadOval 
 .BeginArrowheadWidth = msoArrowheadNarrow 
 .EndArrowheadLength = msoArrowheadLong 
 .EndArrowheadStyle = msoArrowheadTriangle 
 .EndArrowheadWidth = msoArrowheadWide 
End With 

```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]