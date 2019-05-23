---
title: WorksheetFunction.ISO_Ceiling method (Excel)
keywords: vbaxl10.chm137393
f1_keywords:
- vbaxl10.chm137393
ms.prod: excel
api_name:
- Excel.WorksheetFunction.ISO_Ceiling
ms.assetid: e7011c98-0165-a333-6b99-b455913e8575
ms.date: 05/23/2019
localization_priority: Normal
---


# WorksheetFunction.ISO_Ceiling method (Excel)

Returns a number that is rounded up to the nearest integer or to the nearest multiple of significance.


## Syntax

_expression_.**ISO_Ceiling** (_Arg1_, _Arg2_)

_expression_ A variable that represents a **[WorksheetFunction](Excel.WorksheetFunction.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Double**|Number - The value to be rounded.|
| _Arg2_|Optional| **Variant**|Significance - The optional multiple to which number is to be rounded. If significance is omitted, its default value is 1.<br/><br/>**NOTE**: The absolute value of the multiple is used, so that the **ISO_Ceiling** function returns the mathematical ceiling irrespective of the signs of number and significance.|


## Return value

**Double**




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]