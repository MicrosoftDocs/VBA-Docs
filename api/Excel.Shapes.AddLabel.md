---
title: Shapes.AddLabel method (Excel)
keywords: vbaxl10.chm638080
f1_keywords:
- vbaxl10.chm638080
ms.prod: excel
api_name:
- Excel.Shapes.AddLabel
ms.assetid: eb0bfb2a-51ab-ce65-0ef2-aa964d3b08ba
ms.date: 05/15/2019
localization_priority: Normal
---


# Shapes.AddLabel method (Excel)

Creates a label. Returns a **[Shape](Excel.Shape.md)** object that represents the new label.


## Syntax

_expression_.**AddLabel** (_Orientation_, _Left_, _Top_, _Width_, _Height_)

_expression_ A variable that represents a **[Shapes](Excel.Shapes.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Orientation_|Required| **[MsoTextOrientation](Office.MsoTextOrientation.md)**|The text orientation within the label.|
| _Left_|Required| **Single**|The position (in [points](../language/glossary/vbe-glossary.md#point)) of the upper-left corner of the label relative to the upper-left corner of the document.|
| _Top_|Required| **Single**|The position (in points) of the upper-left corner of the label relative to the top of the document.|
| _Width_|Required| **Single**|The width of the label, in points.|
| _Height_|Required| **Single**|The height of the label, in points.|

## Return value

**Shape**


## Example

This example adds a vertical label that contains the text Test Label to _myDocument_.

```vb
Set myDocument = Worksheets(1) 
myDocument.Shapes.AddLabel(msoTextOrientationVertical, _ 
    100, 100, 60, 150) _ 
    .TextFrame.Characters.Text = "Test Label"
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]