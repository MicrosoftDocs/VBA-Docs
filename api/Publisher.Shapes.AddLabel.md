---
title: Shapes.AddLabel method (Publisher)
keywords: vbapb10.chm2162707
f1_keywords:
- vbapb10.chm2162707
ms.prod: publisher
api_name:
- Publisher.Shapes.AddLabel
ms.assetid: 5a803aa2-d37f-6da1-7d8b-58ee2dcd8146
ms.date: 06/14/2019
localization_priority: Normal
---


# Shapes.AddLabel method (Publisher)

Adds a new **[Shape](Publisher.Shape.md)** object representing a text label to the specified **Shapes** collection.


## Syntax

_expression_.**AddLabel** (_Orientation_, _Left_, _Top_, _Width_, _Height_)

_expression_ A variable that represents a **[Shapes](Publisher.Shapes.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_Orientation_|Required| **[PbTextOrientation](publisher.pbtextorientation.md)**|The orientation of the label. Can be one of the **PbTextOrientation** constants.|
|_Left_ |Required| **Variant**|The position of the left edge of the shape representing the text label.|
|_Top_ |Required| **Variant**|The position of the top edge of the shape representing the text label.|
|_Width_|Required| **Variant**|The width of the shape representing the text label.|
|_Height_|Required| **Variant**|The height of the shape representing the text label.|

## Return value

Shape


## Remarks

For the _Left_, _Top_, _Width_, and _Height_ arguments, numeric values are evaluated in [points](../language/glossary/vbe-glossary.md#point); strings can be in any units supported by Microsoft Publisher (for example, "2.5 in").


## Example

The following example adds a new horizontal text label to the first page of the active publication.

```vb
Dim shpLabel As Shape 
 
Set shpLabel = ActiveDocument.Pages(1).Shapes.AddLabel _ 
 (Orientation:=pbTextOrientationHorizontal, _ 
 Left:=144, Top:=144, _ 
 Width:=72, Height:=18)
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]