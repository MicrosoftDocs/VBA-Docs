---
title: Application.PointsToLines method (Publisher)
keywords: vbapb10.chm131158
f1_keywords:
- vbapb10.chm131158
ms.prod: publisher
api_name:
- Publisher.Application.PointsToLines
ms.assetid: beab39fe-9458-6878-ae45-487a8b2271df
ms.date: 06/05/2019
localization_priority: Normal
---


# Application.PointsToLines method (Publisher)

Converts a measurement from [points](../language/glossary/vbe-glossary.md#point) to lines (1 line = 12 points). Returns the converted measurement as a **Single**.


## Syntax

_expression_.**PointsToLines** (_Value_)

_expression_ A variable that represents an **[Application](Publisher.Application.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_Value_|Required| **Single**|The point value to be converted to lines.|

## Return value

Single


## Remarks

This method assumes a measurement in 12-point lines; the actual size of any text in the publication has no effect on the conversion factor.

Use the **[LinesToPoints](Publisher.Application.LinesToPoints.md)** method to convert measurements in lines to points.


## Example

This example converts measurements in points to measurements in lines, demonstrating that the font size in the current selection has no bearing on the conversion factor. Some text must be selected in the active publication for this example to work.

```vb
Dim strOutput As String 
 
' Set text size to 10 points. 
Selection.TextRange.Font.Size = 10 
 
' Display result for 12 points. 
strOutput = "12 points = " _ 
 & Format(Application _ 
 .PointsToLines(Value:=12), _ 
 "0.00") & " lines"
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]