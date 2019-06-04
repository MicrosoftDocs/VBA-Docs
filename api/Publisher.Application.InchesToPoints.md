---
title: Application.InchesToPoints method (Publisher)
keywords: vbapb10.chm131143
f1_keywords:
- vbapb10.chm131143
ms.prod: publisher
api_name:
- Publisher.Application.InchesToPoints
ms.assetid: 32c8740f-ad14-c947-b960-500378a5873d
ms.date: 06/05/2019
localization_priority: Normal
---


# Application.InchesToPoints method (Publisher)

Converts a measurement from inches to [points](../language/glossary/vbe-glossary.md#point) (1 inch = 72 points). Returns the converted measurement as a **Single**.


## Syntax

_expression_.**InchesToPoints** (_Value_)

_expression_ A variable that represents an **[Application](Publisher.Application.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_Value_|Required| **Single**|The inch value to be converted to points.|

## Return value

Single


## Remarks

Use the **[PointsToInches](Publisher.Application.PointsToInches.md)** method to convert measurements in points to inches.


## Example

This example converts measurements in inches entered by the user to measurements in points.

```vb
Dim strInput As String 
Dim strOutput As String 
 
Do While True 
 ' Get input from user. 
 strInput = InputBox( _ 
 Prompt:="Enter measurement in inches (0 to cancel): ", _ 
 Default:="0") 
 
 ' Exit the loop if user enters zero. 
 If Val(strInput) = 0 Then Exit Do 
 
 ' Evaluate and display result. 
 strOutput = Trim(strInput) & " in = " _ 
 & Format(Application _ 
 .InchesToPoints(Value:=Val(strInput)), _ 
 "0.00") & " points" 
 
 MsgBox strOutput 
Loop 

```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]