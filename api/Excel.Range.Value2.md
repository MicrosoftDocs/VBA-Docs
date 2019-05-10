---
title: Range.Value2 property (Excel)
keywords: vbaxl10.chm144217
f1_keywords:
- vbaxl10.chm144217
ms.prod: excel
api_name:
- Excel.Range.Value2
ms.assetid: 0a5d7e6f-2886-5048-66ad-a5078e3465e7
ms.date: 05/11/2019
localization_priority: Normal
---


# Range.Value2 property (Excel)

Returns or sets the cell value. Read/write **Variant**.


## Syntax

_expression_.**Value2**

_expression_ A variable that represents a **[Range](excel.range(object).md)** object.


## Remarks

The only difference between this property and the **Value** property is that the **Value2** property doesn't use the **Currency** and **Date** data types. You can return values formatted with these data types as floating-point numbers by using the **Double** data type.


## Example

This example uses the **Value2** property to add the values of two cells.

```vb
Range("a1").Value2 = Range("b1").Value2 + Range("c1").Value2
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
