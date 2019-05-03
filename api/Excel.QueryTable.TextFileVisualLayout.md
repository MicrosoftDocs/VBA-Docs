---
title: QueryTable.TextFileVisualLayout property (Excel)
keywords: vbaxl10.chm518137
f1_keywords:
- vbaxl10.chm518137
ms.prod: excel
api_name:
- Excel.QueryTable.TextFileVisualLayout
ms.assetid: 13105ba8-945d-9e9b-f90c-9059e2ade9f1
ms.date: 05/03/2019
localization_priority: Normal
---


# QueryTable.TextFileVisualLayout property (Excel)

Returns or sets an **[XlTextVisualLayoutType](Excel.XlTextVisualLayoutType.md)** enumeration that indicates whether the visual layout of the text being imported is left-to-right or right-to-left.


## Syntax

_expression_.**TextFileVisualLayout**

_expression_ A variable that represents a **[QueryTable](Excel.QueryTable.md)** object.


## Remarks

If you import data by using the user interface, data from a web query or a text query is imported as a **QueryTable** object, while all other external data is imported as a **[ListObject](Excel.ListObject.md)** object.

If you import data by using the object model, data from a web query or a text query must be imported as a **QueryTable**, while all other external data can be imported as either a **ListObject** or a **QueryTable**.

The **TextFileVisualLayout** property applies only to **QueryTable** objects.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]