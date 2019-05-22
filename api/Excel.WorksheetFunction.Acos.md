---
title: WorksheetFunction.Acos method (Excel)
keywords: vbaxl10.chm137120
f1_keywords:
- vbaxl10.chm137120
ms.prod: excel
api_name:
- Excel.WorksheetFunction.Acos
ms.assetid: 76954fdf-5aa0-de8d-1f7c-4184ebc472f4
ms.date: 05/22/2019
localization_priority: Normal
---


# WorksheetFunction.Acos method (Excel)

Returns the arccosine, or inverse cosine, of a number. The arccosine is the angle whose cosine is _Arg1_. The returned angle is given in radians in the range 0 (zero) to pi.


## Syntax

_expression_.**Acos** (_Arg1_)

_expression_ A variable that represents a **[WorksheetFunction](Excel.WorksheetFunction.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Double**|The cosine of the angle you want, and must be from -1 to 1.|


## Return value

**Double** 


## Remarks

If you want to convert the result from radians to degrees, multiply it by 180/PI() or use the **[Degrees](Excel.WorksheetFunction.Degrees.md)** method.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]