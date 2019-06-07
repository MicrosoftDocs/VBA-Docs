---
title: LineFormat.BeginArrowheadLength property (Publisher)
keywords: vbapb10.chm3408129
f1_keywords:
- vbapb10.chm3408129
ms.prod: publisher
api_name:
- Publisher.LineFormat.BeginArrowheadLength
ms.assetid: 87daaecf-3b2b-7f21-47fd-bdf192dcac60
ms.date: 06/08/2019
localization_priority: Normal
---


# LineFormat.BeginArrowheadLength property (Publisher)

Returns or sets an **[MsoArrowheadLength](Office.MsoArrowheadLength.md)** constant indicating the length of the arrowhead at the beginning of the specified line. Read/write.


## Syntax

_expression_.**BeginArrowheadLength**

_expression_ A variable that represents a **[LineFormat](Publisher.LineFormat.md)** object.


## Return value

MsoArrowheadLength


## Remarks

The **BeginArrowheadLength** property value can be one of the **MsoArrowheadLength** constants declared in the Microsoft Office type library.

Use the **[EndArrowheadLength](Publisher.LineFormat.EndArrowheadLength.md)** property to return or set the length of the arrowhead at the end of the line.


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