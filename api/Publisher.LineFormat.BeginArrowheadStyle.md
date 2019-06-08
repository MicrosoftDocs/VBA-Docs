---
title: LineFormat.BeginArrowheadStyle property (Publisher)
keywords: vbapb10.chm3408130
f1_keywords:
- vbapb10.chm3408130
ms.prod: publisher
api_name:
- Publisher.LineFormat.BeginArrowheadStyle
ms.assetid: 93dcf2ed-07a3-4391-dd46-2ff9cf89ef36
ms.date: 06/08/2019
localization_priority: Normal
---


# LineFormat.BeginArrowheadStyle property (Publisher)

Returns or sets an **[MsoArrowheadStyle](Office.MsoArrowheadStyle.md)** constant indicating the style of the arrowhead at the beginning of the specified line. Read/write.


## Syntax

_expression_.**BeginArrowheadStyle**

_expression_ A variable that represents a **[LineFormat](Publisher.LineFormat.md)** object.


## Return value

MsoArrowheadStyle


## Remarks

The **BeginArrowheadStyle** property value can be one of the **MsoArrowheadStyle** constants declared in the Microsoft Office type library.

Use the **[EndArrowheadStyle](Publisher.LineFormat.EndArrowheadStyle.md)** property to return or set the style of the arrowhead at the end of the line.


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