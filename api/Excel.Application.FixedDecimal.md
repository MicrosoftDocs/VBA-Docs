---
title: Application.FixedDecimal property (Excel)
keywords: vbaxl10.chm133138
f1_keywords:
- vbaxl10.chm133138
api_name:
- Excel.Application.FixedDecimal
ms.assetid: 49b0a3de-bf5a-0130-e473-5b52f761932a
ms.date: 04/04/2019
ms.localizationpriority: medium
---


# Application.FixedDecimal property (Excel)

All data entered after this property is set to **True** will be formatted with the number of fixed decimal places set by the **[FixedDecimalPlaces](Excel.Application.FixedDecimalPlaces.md)** property. Read/write **Boolean**.


## Syntax

_expression_.**FixedDecimal**

_expression_ A variable that represents an **[Application](Excel.Application(object).md)** object.


## Example

This example sets the **FixedDecimal** property to **True** and then sets the **FixedDecimalPlaces** property to 4. Entering "30000" after running this example produces "3" on the worksheet, and entering "12500" produces "1.25."

```vb
Application.FixedDecimal = True 
Application.FixedDecimalPlaces = 4
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]