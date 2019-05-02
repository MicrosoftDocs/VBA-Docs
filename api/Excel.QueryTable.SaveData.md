---
title: QueryTable.SaveData property (Excel)
keywords: vbaxl10.chm518095
f1_keywords:
- vbaxl10.chm518095
ms.prod: excel
api_name:
- Excel.QueryTable.SaveData
ms.assetid: 7657e1ee-cbed-91c6-0e69-defe4ca69897
ms.date: 05/03/2019
localization_priority: Normal
---


# QueryTable.SaveData property (Excel)

**True** if data for the QueryTable report is saved with the workbook. **False** if only the report definition is saved. Read/write **Boolean**.


## Syntax

_expression_.**SaveData**

_expression_ A variable that represents a **[QueryTable](Excel.QueryTable.md)** object.


## Remarks

For OLAP data sources, this property is always set to **False**.

If you import data by using the user interface, data from a web query or a text query is imported as a **QueryTable** object, while all other external data is imported as a **[ListObject](Excel.ListObject.md)** object.

If you import data by using the object model, data from a web query or a text query must be imported as a **QueryTable**, while all other external data can be imported as either a **ListObject** or a **QueryTable**.

You can use the **[QueryTable](Excel.ListObject.QueryTable.md)** property of the **ListObject** to access the **SaveData** property.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]