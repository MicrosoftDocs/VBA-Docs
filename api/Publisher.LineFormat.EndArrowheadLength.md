---
title: LineFormat.EndArrowheadLength Property (Publisher)
keywords: vbapb10.chm3408133
f1_keywords:
- vbapb10.chm3408133
ms.prod: publisher
api_name:
- Publisher.LineFormat.EndArrowheadLength
ms.assetid: 3e46e63b-54b2-edbf-0dc1-fba2c3a5d945
ms.date: 06/08/2017
---


# LineFormat.EndArrowheadLength Property (Publisher)

Returns or sets an  **MsoArrowheadLength** constant indicating the length of the arrowhead at the end of the specified line. Read/write.


## Syntax

 _expression_. **EndArrowheadLength**

 _expression_ A variable that represents an  **LineFormat** object.


### Return value

MsoArrowheadLength


## Remarks

Use the  **[BeginArrowheadLength](Publisher.LineFormat.BeginArrowheadLength.md)** property to return or set the length of the arrowhead at the beginning of the line.

The  **EndArrowheadLength** property value can be one of the ** [MsoArrowheadLength](./Office.MsoArrowheadLength.md)** constants declared in the Microsoft Office type library.


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


