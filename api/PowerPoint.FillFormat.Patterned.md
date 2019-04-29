---
title: FillFormat.Patterned method (PowerPoint)
keywords: vbapp10.chm552004
f1_keywords:
- vbapp10.chm552004
ms.prod: powerpoint
api_name:
- PowerPoint.FillFormat.Patterned
ms.assetid: 665c5b1d-e2a2-64ab-a0c3-7d22d8d3121a
ms.date: 06/08/2017
localization_priority: Normal
---


# FillFormat.Patterned method (PowerPoint)

Sets the specified fill to a pattern.


## Syntax

_expression_.**Patterned** (_Pattern_)

_expression_ A variable that represents a **[FillFormat](powerpoint.fillformat.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Pattern_|Required|**MsoPatternType**|The pattern to be used for the specified fill. See Remarks for possible values.|

## Remarks

Use the [BackColor](PowerPoint.FillFormat.BackColor.md)and  **[ForeColor](PowerPoint.FillFormat.ForeColor.md)** properties to set the colors used in the pattern.

The value of the Pattern parameter can be one of these  **MsoPatternType** constants.


||
|:-----|
|**msoPattern10Percent**|
|**msoPattern20Percent**|
|**msoPattern25Percent**|
|**msoPattern30Percent**|
|**msoPattern40Percent**|
|**msoPattern50Percent**|
|**msoPattern5Percent**|
|**msoPattern60Percent**|
|**msoPattern70Percent**|
|**msoPattern75Percent**|
|**msoPattern80Percent**|
|**msoPattern90Percent**|
|**msoPatternDarkDownwardDiagonal**|
|**msoPatternDarkHorizontal**|
|**msoPatternDarkUpwardDiagonal**|
|**msoPatternDashedDownwardDiagonal**|
|**msoPatternDashedHorizontal**|
|**msoPatternDashedUpwardDiagonal**|
|**msoPatternDashedVertical**|
|**msoPatternDiagonalBrick**|
|**msoPatternDivot**|
|**msoPatternDottedDiamond**|
|**msoPatternDottedGrid**|
|**msoPatternHorizontalBrick**|
|**msoPatternLargeCheckerBoard**|
|**msoPatternLargeConfetti**|
|**msoPatternLargeGrid**|
|**msoPatternLightDownwardDiagonal**|
|**msoPatternLightHorizontal**|
|**msoPatternLightUpwardDiagonal**|
|**msoPatternLightVertical**|
|**msoPatternMixed**|
|**msoPatternNarrowHorizontal**|
|**msoPatternNarrowVertical**|
|**msoPatternOutlinedDiamond**|
|**msoPatternPlaid**|
|**msoPatternShingle**|
|**msoPatternSmallCheckerBoard**|
|**msoPatternSmallConfetti**|
|**msoPatternSmallGrid**|
|**msoPatternSolidDiamond**|
|**msoPatternSphere**|
|**msoPatternTrellis**|
|**msoPatternWave**|
|**msoPatternWeave**|
|**msoPatternWideDownwardDiagonal**|
|**msoPatternWideUpwardDiagonal**|
|**msoPatternZigZag**|
|**msoPatternDarkVertical**|

## Example

This example adds an oval with a patterned fill to _myDocument_.


```vb
Set myDocument = ActivePresentation.Slides(1)

With myDocument.Shapes.AddShape(msoShapeOval, 60, 60, 80, 40).Fill

    .ForeColor.RGB = RGB(128, 0, 0)

    .BackColor.RGB = RGB(0, 0, 255)

    .Patterned msoPatternDarkVertical

End With
```


## See also


[FillFormat Object](PowerPoint.FillFormat.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]