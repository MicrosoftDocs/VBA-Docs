---
title: Application.PicasToPoints method (Publisher)
keywords: vbapb10.chm131152
f1_keywords:
- vbapb10.chm131152
ms.prod: publisher
api_name:
- Publisher.Application.PicasToPoints
ms.assetid: 64d3e435-dcc1-d637-7aac-cc9a9bf81e76
ms.date: 06/05/2019
localization_priority: Normal
---


# Application.PicasToPoints method (Publisher)

Converts a measurement from picas to [points](../language/glossary/vbe-glossary.md#point) (1 pica = 12 points). Returns the converted measurement as a **Single**.


## Syntax

_expression_.**PicasToPoints** (_Value_)

_expression_ A variable that represents an **[Application](Publisher.Application.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_Value_|Required| **Single**|The pica value to be converted to points.|

## Return value

Single


## Remarks

Use the **[PointsToPicas](Publisher.Application.PointsToPicas.md)** method to convert measurements in points to picas.


## Example

This example converts measurements in picas entered by the user to measurements in points.

```vb
Dim strInput As String 
Dim strOutput As String 
 
Do While True 
 ' Get input from user. 
 strInput = InputBox( _ 
 Prompt:="Enter measurement in picas (0 to cancel): ", _ 
 Default:="0") 
 
 ' Exit the loop if user enters zero. 
 If Val(strInput) = 0 Then Exit Do 
 
 ' Evaluate and display result. 
 strOutput = Trim(strInput) & " picas = " _ 
 & Format(Application _ 
 .PicasToPoints(Value:=Val(strInput)), _ 
 "0.00") & " points" 
 
 MsgBox strOutput 
Loop
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]