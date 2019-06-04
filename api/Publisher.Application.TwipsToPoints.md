---
title: Application.TwipsToPoints method (Publisher)
keywords: vbapb10.chm131154
f1_keywords:
- vbapb10.chm131154
ms.prod: publisher
api_name:
- Publisher.Application.TwipsToPoints
ms.assetid: 18e1c4da-1295-31a2-d66b-ab0df807b7a6
ms.date: 06/05/2019
localization_priority: Normal
---


# Application.TwipsToPoints method (Publisher)

Converts a measurement from [twips](../language/glossary/vbe-glossary.md#twip) to [points](../language/glossary/vbe-glossary.md#point) (20 twips = 1 point). Returns the converted measurement as a **Single**.


## Syntax

_expression_.**TwipsToPoints** (_Value_)

_expression_ A variable that represents an **[Application](Publisher.Application.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_Value_|Required| **Single**|The twip value to be converted to points.|

## Return value

Single


## Remarks

Use the **[PointsToTwips](Publisher.Application.PointsToTwips.md)** method to convert measurements in points to twips.


## Example

This example converts measurements in twips entered by the user to measurements in points.

```vb
Dim strInput As String 
Dim strOutput As String 
 
Do While True 
 ' Get input from user. 
 strInput = InputBox( _ 
 Prompt:="Enter measurement in twips (0 to cancel): ", _ 
 Default:="0") 
 
 ' Exit the loop if user enters zero. 
 If Val(strInput) = 0 Then Exit Do 
 
 ' Evaluate and display result. 
 strOutput = Trim(strInput) & " twips = " _ 
 & Format(Application _ 
 .TwipsToPoints(Value:=Val(strInput)), _ 
 "0.00") & " points" 
 
 MsgBox strOutput 
Loop 

```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]