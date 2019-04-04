---
title: Application.CalculateFull method (Excel)
keywords: vbaxl10.chm133255
f1_keywords:
- vbaxl10.chm133255
ms.prod: excel
api_name:
- Excel.Application.CalculateFull
ms.assetid: 11be6386-a5de-817f-0624-b7e7fd502ed3
ms.date: 04/04/2019
localization_priority: Normal
---


# Application.CalculateFull method (Excel)

Forces a full calculation of the data in all open workbooks.


## Syntax

_expression_.**CalculateFull**

_expression_ A variable that represents an **[Application](Excel.Application(object).md)** object.


## Example

This example compares the version of Microsoft Excel with the version of Excel that the workbook was last calculated in. If the two version numbers are different, a full calculation of the data in all open workbooks is performed.

```vb
If Application.CalculationVersion <> _ 
 Workbooks(1).CalculationVersion Then 
 Application.CalculateFull 
End If
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
