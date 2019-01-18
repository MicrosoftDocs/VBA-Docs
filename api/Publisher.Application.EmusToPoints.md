---
title: Application.EmusToPoints Method (Publisher)
keywords: vbapb10.chm131142
f1_keywords:
- vbapb10.chm131142
ms.prod: publisher
api_name:
- Publisher.Application.EmusToPoints
ms.assetid: 941e5975-ca7a-38dc-8116-e90b2a2ab6e5
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.EmusToPoints Method (Publisher)

Converts a measurement from emus to points (12700 emus = 1 point). Returns the converted measurement as a  **Single**.


## Syntax

 _expression_. **EmusToPoints**(**_Value_**)

 _expression_ A variable that represents an  **Application** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|Value|Required| **Single**|An expression that returns one of the objects in the **Applies To** list.|

## Return value

Single


## Remarks

Use the  **[PointsToEmus](Publisher.Application.PointsToEmus.md)** method to convert measurements in points to emus.


## Example

This example converts measurements in emus entered by the user to measurements in points.


```vb
Dim strInput As String 
Dim strOutput As String 
 
Do While True 
 ' Get input from user. 
 strInput = InputBox( _ 
 Prompt:="Enter measurement in emus (0 to cancel): ", _ 
 Default:="0") 
 
 ' Exit the loop if user enters zero. 
 If Val(strInput) = 0 Then Exit Do 
 
 ' Evaluate and display result. 
 strOutput = Trim(strInput) & " emus = " _ 
 & Format(Application _ 
 .EmusToPoints(Value:=Val(strInput)), _ 
 "0.00") & " points" 
 
 MsgBox strOutput 
Loop 

```


## See also


 [Application Object](Publisher.Application.md)

