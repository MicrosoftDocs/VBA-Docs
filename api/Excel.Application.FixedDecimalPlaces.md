---
title: Application.FixedDecimalPlaces property (Excel)
keywords: vbaxl10.chm133139
f1_keywords:
- vbaxl10.chm133139
ms.prod: excel
api_name:
- Excel.Application.FixedDecimalPlaces
ms.assetid: e264dce3-4589-3e83-c931-5d69e3b8b3be
ms.date: 04/04/2019
localization_priority: Normal
---


# Application.FixedDecimalPlaces property (Excel)

Returns or sets the number of fixed decimal places used when the **[FixedDecimal](Excel.Application.FixedDecimal.md)** property is set to **True**. Read/write **Long**.


## Syntax

_expression_.**FixedDecimalPlaces**

_expression_ A variable that represents an **[Application](Excel.Application(object).md)** object.


## Example

This example sets the **FixedDecimal** property to **True** and then sets the **FixedDecimalPlaces** property to 4. Entering "30000" after running this example produces "3" on the worksheet, and entering "12500" produces "1.25."

```vb
Application.FixedDecimal = True 
Application.FixedDecimalPlaces = 4
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]