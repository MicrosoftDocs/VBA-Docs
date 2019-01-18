---
title: Application.PointsToMillimeters Method (Publisher)
keywords: vbapb10.chm131159
f1_keywords:
- vbapb10.chm131159
ms.prod: publisher
api_name:
- Publisher.Application.PointsToMillimeters
ms.assetid: eaa9154d-1a9b-81e7-58bc-3f7bf873ab97
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.PointsToMillimeters Method (Publisher)

Converts a measurement from points to millimeters (1 mm = 2.835 points). Returns the converted measurement as a  **Single**.


## Syntax

 _expression_. **PointsToMillimeters**(**_Value_**)

 _expression_ A variable that represents an  **Application** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|Value|Required| **Single**|The point value to be converted to millimeters.|

## Return value

Single


## Remarks

Use the  **[MillimetersToPoints](Publisher.Application.MillimetersToPoints.md)** method to convert measurements in millimeters to points.


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
 .PointsToMillimeters(Value:=Val(strInput)), _ 
 "0.00") & " mm" 
 
 MsgBox strOutput 
Loop 

```


## See also


 [Application Object](Publisher.Application.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]