---
title: Application.PointsToEmus method (Publisher)
keywords: vbapb10.chm131156
f1_keywords:
- vbapb10.chm131156
ms.prod: publisher
api_name:
- Publisher.Application.PointsToEmus
ms.assetid: cb3f0bb9-fa0d-d967-9294-081a369c2c4e
ms.date: 06/05/2019
localization_priority: Normal
---


# Application.PointsToEmus method (Publisher)

Converts a measurement from [points](../language/glossary/vbe-glossary.md#point) to emus (12700 emus = 1 point). Returns the converted measurement as a **Single**.


## Syntax

_expression_.**PointsToEmus** (_Value_)

_expression_ A variable that represents an **[Application](Publisher.Application.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_Value_|Required| **Single**|The point value to be converted to emus.|

## Return value

Single


## Remarks

Use the **[EmusToPoints](Publisher.Application.EmusToPoints.md)** method to convert measurements in emus to points.


## Example

This example converts measurements in points entered by the user to measurements in emus.

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
 .PointsToEmus(Value:=Val(strInput)), _ 
 "0.00") & " emus" 
 
 MsgBox strOutput 
Loop 

```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]