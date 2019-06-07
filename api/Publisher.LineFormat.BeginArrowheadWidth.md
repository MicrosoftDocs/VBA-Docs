---
title: LineFormat.BeginArrowheadWidth property (Publisher)
keywords: vbapb10.chm3408131
f1_keywords:
- vbapb10.chm3408131
ms.prod: publisher
api_name:
- Publisher.LineFormat.BeginArrowheadWidth
ms.assetid: a752c674-1b83-b8c8-d325-b61804f5fadc
ms.date: 06/08/2019
localization_priority: Normal
---


# LineFormat.BeginArrowheadWidth property (Publisher)

Returns or sets an **[MsoArrowheadWidth](Office.MsoArrowheadWidth.md)** constant indicating the width of the arrowhead at the beginning of the specified line. Read/write.


## Syntax

_expression_.**BeginArrowheadWidth**

_expression_ A variable that represents a **[LineFormat](Publisher.LineFormat.md)** object.


## Return value

MsoArrowheadWidth


## Remarks

The **BeginArrowheadWidth** property value can be one of the **MsoArrowheadWidth** constants declared in the Microsoft Office type library.

Use the **[EndArrowheadWidth](Publisher.LineFormat.EndArrowheadWidth.md)** property to return or set the width of the arrowhead at the end of the line.


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