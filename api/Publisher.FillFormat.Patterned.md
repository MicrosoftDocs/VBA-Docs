---
title: FillFormat.Patterned method (Publisher)
keywords: vbapb10.chm2359314
f1_keywords:
- vbapb10.chm2359314
ms.prod: publisher
api_name:
- Publisher.FillFormat.Patterned
ms.assetid: 10e363b7-1160-55d3-5c97-733b7742b619
ms.date: 06/07/2019
localization_priority: Normal
---


# FillFormat.Patterned method (Publisher)

Sets the specified fill to a pattern.


## Syntax

_expression_.**Patterned** (_Pattern_)

_expression_ A variable that represents a **[FillFormat](publisher.fillformat.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_Pattern_|Required| **[MsoPatternType](Office.MsoPatternType.md)** |The pattern to be used for the specified fill. Can be one of the **MsoPatternType** constants declared in the Microsoft Office type library.|

## Remarks

Use the **[BackColor](Publisher.FillFormat.BackColor.md)** and **[ForeColor](Publisher.FillFormat.ForeColor.md)** properties to set the colors used in the pattern.


## Example

This example adds an oval with a patterned fill to the active publication.

```vb
With ActiveDocument.Pages(1).Shapes _ 
 .AddShape(Type:=msoShapeOval, _ 
 Left:=60, Top:=60, Width:=80, Height:=40).Fill 
 .ForeColor.RGB = RGB(128, 0, 0) 
 .BackColor.RGB = RGB(0, 0, 255) 
 .Patterned Pattern:=msoPatternDarkVertical 
End With 

```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]