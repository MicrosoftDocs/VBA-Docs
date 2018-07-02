---
title: LineFormat.EndArrowheadWidth Property (Publisher)
keywords: vbapb10.chm3408135
f1_keywords:
- vbapb10.chm3408135
ms.prod: publisher
api_name:
- Publisher.LineFormat.EndArrowheadWidth
ms.assetid: 20284d2d-e733-ee26-3c1c-53fd60012a75
ms.date: 06/08/2017
---


# LineFormat.EndArrowheadWidth Property (Publisher)

Returns or sets an  **MsoArrowheadWidth** constant indicating the width of the arrowhead at the end of the specified line. Read/write.


## Syntax

 _expression_. **EndArrowheadWidth**

 _expression_ A variable that represents an  **LineFormat** object.


### Return Value

MsoArrowheadWidth


## Remarks

Use the  **[BeginArrowheadWidth](Publisher.LineFormat.BeginArrowheadWidth.md)** property to return or set the width of the arrowhead at the beginning of the line.

The  **EndArrowheadWidth** property value can be one of the ** [MsoArrowheadWidth](./Office.MsoArrowheadWidth.md)** constants declared in the Microsoft Office type library.


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


