---
title: Application.PointsToTwips method (Publisher)
keywords: vbapb10.chm131168
f1_keywords:
- vbapb10.chm131168
ms.prod: publisher
api_name:
- Publisher.Application.PointsToTwips
ms.assetid: ba928b83-f551-049e-5868-098a9837ee7b
ms.date: 06/05/2019
localization_priority: Normal
---


# Application.PointsToTwips method (Publisher)

Converts a measurement from [points](../language/glossary/vbe-glossary.md#point) to [twips](../language/glossary/vbe-glossary.md#twip) (20 twips = 1 point). Returns the converted measurement as a **Single**.


## Syntax

_expression_.**PointsToTwips** (_Value_)

_expression_ A variable that represents an **[Application](Publisher.Application.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_Value_|Required| **Single**|The point value to be converted to twips.|

## Return value

Single


## Remarks

Use the **[TwipsToPoints](Publisher.Application.TwipsToPoints.md)** method to convert measurements in twips to points.


## Example

This example converts measurements in points entered by the user to measurements in twips.

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
 .PointsToTwips(Value:=Val(strInput)), _ 
 "0.00") & " twips" 
 
 MsgBox strOutput 
Loop 

```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]