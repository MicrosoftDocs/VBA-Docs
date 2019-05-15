---
title: Shapes.AddCallout method (Excel)
keywords: vbaxl10.chm638077
f1_keywords:
- vbaxl10.chm638077
ms.prod: excel
api_name:
- Excel.Shapes.AddCallout
ms.assetid: b98ea95d-210b-34cc-c999-e7ce0a3e3a72
ms.date: 05/15/2019
localization_priority: Normal
---


# Shapes.AddCallout method (Excel)

Creates a borderless line callout. Returns a **[Shape](Excel.Shape.md)** object that represents the new callout.


## Syntax

_expression_.**AddCallout** (_Type_, _Left_, _Top_, _Width_, _Height_)

_expression_ A variable that represents a **[Shapes](Excel.Shapes.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Type_|Required| **[MsoCalloutType](Office.MsoCalloutType.md)**|The type of callout line.|
| _Left_|Required| **Single**|The position (in [points](../language/glossary/vbe-glossary.md#point)) of the upper-left corner of the callout's bounding box relative to the upper-left corner of the document.|
| _Top_|Required| **Single**|The position (in points) of the upper-left corner of the callout's bounding box relative to the top of the document.|
| _Width_|Required| **Single**|The width of the callout's bounding box, in points.|
| _Height_|Required| **Single**|The height of the callout's bounding box, in points.|

## Return value

Shape


## Remarks

You can insert a greater variety of callouts by using the **[AddShape](Excel.Shapes.AddShape.md)** method.


## Example

This example adds a borderless callout with a freely rotating one-segment callout line to _myDocument_ and then sets the callout angle to 30 degrees.

```vb
Set myDocument = Worksheets(1) 
myDocument.Shapes.AddCallout(Type:=msoCalloutTwo, _ 
    Left:=50, Top:=50, Width:=200, Height:=100) _ 
    .Callout.Angle = msoCalloutAngle30
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]