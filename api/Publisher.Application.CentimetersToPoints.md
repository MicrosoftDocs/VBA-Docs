---
title: Application.CentimetersToPoints method (Publisher)
keywords: vbapb10.chm131141
f1_keywords:
- vbapb10.chm131141
ms.prod: publisher
api_name:
- Publisher.Application.CentimetersToPoints
ms.assetid: 6eda6692-ea9a-c4ad-6991-066fdc23bd2c
ms.date: 06/04/2019
localization_priority: Normal
---


# Application.CentimetersToPoints method (Publisher)

Converts a measurement from centimeters to [points](../language/glossary/vbe-glossary.md#point) (1 cm = 28.35 points). Returns the converted measurement as a **Single**.


## Syntax

_expression_.**CentimetersToPoints** (_Value_)

_expression_ A variable that represents an **[Application](Publisher.Application.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_Value_|Required| **Single**|The centimeter value to be converted to points.|

## Return value

Single


## Remarks

Use the **[PointsToCentimeters](Publisher.Application.PointsToCentimeters.md)** method to convert measurements in points to centimeters.


## Example

This example converts measurements in centimeters entered by the user to measurements in points.

```vb
Dim strInput As String 
Dim strOutput As String 
 
Do While True 
 ' Get input from user. 
 strInput = InputBox( _ 
 Prompt:="Enter measurement in centimeters (0 to cancel): ", _ 
 Default:="0") 
 
 ' Exit the loop if user enters zero. 
 If Val(strInput) = 0 Then Exit Do 
 
 ' Evaluate and display result. 
 strOutput = Trim(strInput) & " cm = " _ 
 & Format(Application _ 
 .CentimetersToPoints(Value:=Val(strInput)), _ 
 "0.00") & " points" 
 
 MsgBox strOutput 
Loop 

```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]