---
title: Application.PointsToPixels method (Publisher)
keywords: vbapb10.chm131161
f1_keywords:
- vbapb10.chm131161
ms.prod: publisher
api_name:
- Publisher.Application.PointsToPixels
ms.assetid: 9c67fcae-6c93-ddae-cbad-75356e5c5084
ms.date: 06/05/2019
localization_priority: Normal
---


# Application.PointsToPixels method (Publisher)

Converts a measurement from [points](../language/glossary/vbe-glossary.md#point) to pixels (1 pixel = 0.75 points). Returns the converted measurement as a **Single**.


## Syntax

_expression_.**PointsToPixels** (_Value_)

_expression_ A variable that represents an **[Application](Publisher.Application.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_Value_|Required| **Single**|The point value to be converted to pixels.|

## Return value

Single


## Remarks

Use the **[PixelsToPoints](Publisher.Application.PixelsToPoints.md)** method to convert measurements in pixels to points.


## Example

This example converts measurements in points entered by the user to measurements in pixels.

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
 .PointsToPixels(Value:=Val(strInput)), _ 
 "0.00") & " pixels" 
 
 MsgBox strOutput 
Loop 

```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]