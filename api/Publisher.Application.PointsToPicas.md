---
title: Application.PointsToPicas method (Publisher)
keywords: vbapb10.chm131160
f1_keywords:
- vbapb10.chm131160
ms.prod: publisher
api_name:
- Publisher.Application.PointsToPicas
ms.assetid: ff566bef-7032-70f7-7880-ff66cfeca88f
ms.date: 06/05/2019
localization_priority: Normal
---


# Application.PointsToPicas method (Publisher)

Converts a measurement from [points](../language/glossary/vbe-glossary.md#point) to picas (1 pica = 12 points). Returns the converted measurement as a **Single**.


## Syntax

_expression_.**PointsToPicas** (_Value_)

_expression_ A variable that represents an **[Application](Publisher.Application.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_Value_|Required| **Single**|The point value to be converted to picas.|

## Return value

Single


## Remarks

Use the **[PicasToPoints](Publisher.Application.PicasToPoints.md)** method to convert measurements in picas to points.


## Example

This example converts measurements in points entered by the user to measurements in picas.

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
 .PointsToPicas(Value:=Val(strInput)), _ 
 "0.00") & " picas" 
 
 MsgBox strOutput 
Loop
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]