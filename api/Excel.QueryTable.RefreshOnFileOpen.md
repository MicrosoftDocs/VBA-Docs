---
title: QueryTable.RefreshOnFileOpen property (Excel)
keywords: vbaxl10.chm518078
f1_keywords:
- vbaxl10.chm518078
ms.prod: excel
api_name:
- Excel.QueryTable.RefreshOnFileOpen
ms.assetid: 25ee4493-1738-66ce-09d3-9e0e83a677b7
ms.date: 05/03/2019
localization_priority: Normal
---


# QueryTable.RefreshOnFileOpen property (Excel)

**True** if the PivotTable cache or query table is automatically updated each time the workbook is opened. The default value is **False**. Read/write **Boolean**.


## Syntax

_expression_.**RefreshOnFileOpen**

_expression_ A variable that represents a **[QueryTable](Excel.QueryTable.md)** object.


## Remarks

Query tables and PivotTable reports are not automatically refreshed when you open the workbook by using the **[Open](Excel.Workbooks.Open.md)** method of the **Workbooks** object in Visual Basic. Use the **[Refresh](Excel.QueryTable.Refresh.md)** method to refresh the data after the workbook is open.

If you import data by using the user interface, data from a web query or a text query is imported as a **QueryTable** object, while all other external data is imported as a **[ListObject](Excel.ListObject.md)** object.

If you import data by using the object model, data from a web query or a text query must be imported as a **QueryTable**, while all other external data can be imported as either a **ListObject** or a **QueryTable**.

You can use the **[QueryTable](Excel.ListObject.QueryTable.md)** property of the **ListObject** to access the **RefreshOnFileOpen** property.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]