---
title: WorksheetFunction.ErfC method (Excel)
keywords: vbaxl10.chm137301
f1_keywords:
- vbaxl10.chm137301
ms.prod: excel
api_name:
- Excel.WorksheetFunction.ErfC
ms.assetid: 7579d8fb-7cad-bb5a-7fb9-0895ef096858
ms.date: 05/22/2019
localization_priority: Normal
---


# WorksheetFunction.ErfC method (Excel)

Returns the complementary **Erf** function integrated between the specified parameter and infinity.

> [!IMPORTANT] 
> This function has been replaced with one or more new functions that may provide improved accuracy and whose names better reflect their usage. This function is still available for compatibility with earlier versions of Excel. However, if backward compatibility is not required, you should consider using the new functions from now on, because they more accurately describe their functionality.
> 
> For more information about the new function, see the **[ErfC_Precise](Excel.WorksheetFunction.ErfC_Precise.md)** method.

## Syntax

_expression_.**ErfC** (_Arg1_, _Arg2_)

_expression_ A variable that represents a **[WorksheetFunction](Excel.WorksheetFunction.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Variant**|The first argument.|
| _Arg2_|Optional| **Variant**|The second argument.|

## Return value

**Double**


## Remarks

If this function is not available, and returns the #NAME? error, you need to install and load the **Analysis ToolPak** add-in.

If the parameter is nonnumeric, **ErfC** returns the #VALUE! error value.

If the parameter is negative, **ErfC** returns the #NUM! error value.


## Example

The following example displays the complementary **Erf** function of 1 (0.1573).

```vb
=ERFC(1)
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]