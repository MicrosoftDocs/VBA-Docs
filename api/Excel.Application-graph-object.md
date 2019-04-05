---
title: Application object (Excel Graph)
keywords: vbagr10.chm3077640
f1_keywords:
- vbagr10.chm3077640
ms.prod: excel
api_name:
- Excel.Application
ms.assetid: 553a0ee2-83da-6d32-f082-15e93e7b0e4d
ms.date: 04/05/2019
localization_priority: Normal
---


# Application object (Excel Graph)

Represents the entire Graph application. The **Application** object represents the top level of the object hierarchy and contains all of the objects, methods, and properties for the application.


## Remarks

Use the **Application** property to return the **Application** object. 


## Example

The following example applies the **DataSheet** property to the **Application** object.

```vb
myChart.Application.DataSheet.Range("A1").Value = 32
```

## Methods

- [Update](Excel.Update.md)

## Properties

- [Application](excel.application-graph-property.md)

## See also

- [Application object (Excel)](excel.application(object).md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]