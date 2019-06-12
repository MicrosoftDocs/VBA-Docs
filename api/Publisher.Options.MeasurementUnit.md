---
title: Options.MeasurementUnit property (Publisher)
keywords: vbapb10.chm1048594
f1_keywords:
- vbapb10.chm1048594
ms.prod: publisher
api_name:
- Publisher.Options.MeasurementUnit
ms.assetid: 49221e4e-c84a-6706-8f9a-3853283ebb18
ms.date: 06/11/2019
localization_priority: Normal
---


# Options.MeasurementUnit property (Publisher)

Returns or sets a **[PbUnitType](publisher.pbunittype.md)** constant representing the standard measurement unit for Microsoft Publisher. Read/write.


## Syntax

_expression_.**MeasurementUnit**

_expression_ A variable that represents an **[Options](Publisher.Options.md)** object.


## Return value

PbUnitType


## Remarks

The **MeasurementUnit** property value can be one of these **PbUnitType** constants declared in the Publisher type library: 

- **pbUnitCM** sets the unit of measurement to centimeters.
- **pbUnitInch** sets the unit of measurement to inches.
- **pbUnitPica** sets the unit of measurement to picas.
- **pbUnitPoint** sets the unit of measurement to [points](../language/glossary/vbe-glossary.md#point).

All other measurement unit constants do not apply to this property; if used, they return an error.

## Example

This example sets the standard measurement unit for Publisher to points.

```vb
Sub SetUnitOfMeasurement() 
 Options.MeasurementUnit = pbUnitPoint 
End Sub
```

<br/>

This example displays the current unit of measurement.

```vb
Sub GetUnitOfMeasurement() 
 Dim measUnit As PbUnitType 
 Dim strUnit As String 
 
 measUnit = Options.MeasurementUnit 
 
 Select Case measUnit 
 Case 0 
 strUnit = "inches" 
 Case 1 
 strUnit = "centimeters" 
 Case 2 
 strUnit = "picas" 
 Case 3 
 strUnit = "points" 
 End Select 
 
 MsgBox "The current unit of measurement is " & strUnit 
 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]