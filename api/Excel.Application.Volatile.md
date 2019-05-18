---
title: Application.Volatile method (Excel)
keywords: vbaxl10.chm133230
f1_keywords:
- vbaxl10.chm133230
ms.prod: excel
api_name:
- Excel.Application.Volatile
ms.assetid: 27047561-9d76-b37d-100d-1c58e6edf494
ms.date: 04/05/2019
localization_priority: Normal
---


# Application.Volatile method (Excel)

Marks a user-defined function as volatile. A volatile function must be recalculated whenever calculation occurs in any cells on the worksheet. A nonvolatile function is recalculated only when the input variables change. This method has no effect if it's not inside a user-defined function used to calculate a worksheet cell.


## Syntax

_expression_.**Volatile** (_Volatile_)

_expression_ A variable that represents an **[Application](Excel.Application(object).md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Volatile_|Optional| **Variant**| **True** to mark the function as volatile. **False** to mark the function as nonvolatile. The default value is **True**.|

## Example

This example marks the user-defined function _My\_Func_ as volatile. The function will be recalculated when any cell in any workbook in the application window changes value. Recalculation of the function is not restricted to changes or calculation cycles on the worksheet for which this function applies. Therefore, use it moderately to avoid calculation lag.

```vb
Function My_Func() 
 Application.Volatile 
 ' 
 ' Remainder of the function 
 ' 
End Function
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
