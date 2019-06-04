---
title: Application.PointsToCentimeters method (Publisher)
keywords: vbapb10.chm131155
f1_keywords:
- vbapb10.chm131155
ms.prod: publisher
api_name:
- Publisher.Application.PointsToCentimeters
ms.assetid: 9a734d3d-78d2-1e27-63b3-2ad1074e16c1
ms.date: 06/05/2019
localization_priority: Normal
---


# Application.PointsToCentimeters method (Publisher)

Converts a measurement from [points](../language/glossary/vbe-glossary.md#point) to centimeters (1 cm = 28.35 points). Returns the converted measurement as a **Single**.


## Syntax

_expression_.**PointsToCentimeters** (_Value_)

_expression_ A variable that represents an **[Application](Publisher.Application.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_Value_|Required| **Single**|The point value to be converted to centimeters.|

## Return value

Single


## Remarks

Use the **[CentimetersToPoints](Publisher.Application.CentimetersToPoints.md)** method to convert measurements in centimeters to points.


## Example

This example converts measurements in points entered by the user to measurements in centimeters.

```vb
Dim strInput As String 
Dim strOutput As String 
 
Do While True 
 ' Get input from user. 
 strInput = InputBox( _ 
 Prompt:="Enter measurement in points (0 to cancel): ", _ 
 Default:="0") 
 
 ' Exit the loop if user enters zero. 
 If Val(strInput) = 0 Then Exit Do 
 
 ' Evaluate and display result. 
 strOutput = Trim(strInput) & " points = " _ 
 & Format(Application _ 
 .PointsToCentimeters(Value:=Val(strInput)), _ 
 "0.00") & " cm" 
 
 MsgBox strOutput 
Loop
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]