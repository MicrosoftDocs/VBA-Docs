---
title: Adjustments.Item property (Publisher)
keywords: vbapb10.chm2424832
f1_keywords:
- vbapb10.chm2424832
ms.prod: publisher
api_name:
- Publisher.Adjustments.Item
ms.assetid: 9adba87a-d09d-b024-f889-4dcdab961561
ms.date: 06/04/2019
localization_priority: Normal
---


# Adjustments.Item property (Publisher)

Returns or sets a **Variant** indicating the adjustment value specified by the _Index_ argument. Read/write.


## Syntax

_expression_.**Item** (_Index_)

_expression_ A variable that represents an **[Adjustments](Publisher.Adjustments.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_Index_|Required| **Integer**|The index number of the adjustment.|

## Remarks

AutoShapes, connectors, and WordArt objects can have up to eight adjustments.

For linear adjustments, an adjustment value of 0.0 generally corresponds to the left or top edge of the shape, and a value of 1.0 generally corresponds to the right or bottom edge of the shape. However, adjustments can pass beyond shape boundaries for some shapes. 

For radial adjustments, an adjustment value of 1.0 corresponds to the width of the shape. 

For angular adjustments, the adjustment value is specified in degrees.

The **Item** property applies only to shapes that have adjustments.


## Example

This example adds two crosses to the active publication, and then sets the value for adjustment one (the only one for this type of AutoShape) on each cross.

```vb
With ActiveDocument.Pages(1).Shapes 
 .AddShape(Type:=msoShapeCross, Left:=10, Top:=10, Width:=100, _ 
 Height:=100).Adjustments.Item(1) = 0.4 
 .AddShape(Type:=msoShapeCross, Left:=150, Top:=10, Width:=100, _ 
 Height:=100).Adjustments.Item(1) = 0.2 
End With
```

<br/>

This example has the same result as the previous example even though it doesn't explicitly use the **Item** property.

```vb
With ActiveDocument.Pages(1).Shapes 
 .AddShape(Type:=msoShapeCross, Left:=10, Top:=10, Width:=100, _ 
 Height:=100).Adjustments(1) = 0.4 
 .AddShape(Type:=msoShapeCross, Left:=150, Top:=10, Width:=100, _ 
 Height:=100).Adjustments(1) = 0.2 
End With
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]