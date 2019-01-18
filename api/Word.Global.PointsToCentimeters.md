---
title: Global.PointsToCentimeters method (Word)
keywords: vbawd10.chm163119485
f1_keywords:
- vbawd10.chm163119485
ms.prod: word
api_name:
- Word.Global.PointsToCentimeters
ms.assetid: 82f64aa2-3c9c-a3e0-5326-8d4227cee170
ms.date: 06/08/2017
localization_priority: Normal
---


# Global.PointsToCentimeters method (Word)

Converts a measurement from points to centimeters (1 centimeter = 28.35 points). Returns the converted measurement as a  **Single**.


## Syntax

 _expression_. `PointsToCentimeters`( `_Points_` )

 _expression_ A variable that represents a '[Global](Word.Global.md)' object. Optional.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Points_|Required| **Single**|The measurement, in points.|

## Return value

Single


## Example

This example converts a measurement of 30 points to the corresponding number of centimeters.


```vb
MsgBox PointsToCentimeters(30) & " centimeters"
```

This example converts the value of the variable  _sngData_ (a measurement in points) to centimeters, inches, lines, millimeters, or picas, depending on the value of the variable _intUnit_ (a value from 1 through 5 that indicates the resulting unit of measurement).




```vb
Function ConvertPoints(ByVal intUnit As Integer, _ 
 sngData As Single) As Single 
 
 Select Case intUnit 
 Case 1 
 ConvertPoints = PointsToCentimeters(sngData) 
 Case 2 
 ConvertPoints = PointsToInches(sngData) 
 Case 3 
 ConvertPoints = PointsToLines(sngData) 
 Case 4 
 ConvertPoints = PointsToMillimeters(sngData) 
 Case 5 
 ConvertPoints = PointsToPicas(sngData) 
 Case Else 
 Error 5 
 End Select 
 
End Function
```


## See also


[Global Object](Word.Global.md)

