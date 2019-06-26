---
title: Cell.ResultForce property (Visio)
keywords: vis_sdr.chm10114200
f1_keywords:
- vis_sdr.chm10114200
ms.prod: visio
api_name:
- Visio.Cell.ResultForce
ms.assetid: 96579953-05f2-edf5-02d6-54ef0e632215
ms.date: 06/08/2017
localization_priority: Normal
---


# Cell.ResultForce property (Visio)

Sets a cell's value, even if the cell's formula is protected with the GUARD function. Read/write.


## Syntax

_expression_.**ResultForce** (_UnitsNameOrCode_)

_expression_ A variable that represents a **[Cell](Visio.Cell.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _UnitsNameOrCode_|Required| **Variant**|The units to use when setting the cell's value.|

## Return value

Double


## Remarks

Use the  **ResultForce** method to set a cell's value even if the cell's formula is protected with a GUARD function. If the string is invalid, an error is generated.

Setting the  **ResultForce** property is similar to setting a cell's **ResultFromIntForce** property. The difference is that the **ResultFromIntForce** property accepts an integer for the value of the cell, whereas the **ResultForce** property accepts a floating point number.

You can specify  _UnitsNameOrCode_ as an integer or a string value. For example, the following statements all set _UnitsNameOrCode_ to inches.

 **Cell.ResultForce** (**visInches**) = _newValue_

 **Cell.ResultForce** (65) = _newValue_

 **Cell.ResultForce** ("in") = _newValue_ where "in" can also be any of the alternate strings representing inches, such as "inch", "in.", or "intCounter".

For a complete list of valid unit strings along with their corresponding Automation constants (integer values), see [About units of measure](../visio/Concepts/about-units-of-measure-visio.md).

Automation constants for representing units are declared by the Visio type library in member  **[VisUnitCodes](Visio.visunitcodes.md)**.

To specify internal units, pass a zero-length string (""). Internal units are inches for distance and radians for angles. To specify implicit units, you must use the  **Formula** property.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]