---
title: QueryTable.WebTables property (Excel)
keywords: vbaxl10.chm518124
f1_keywords:
- vbaxl10.chm518124
ms.prod: excel
api_name:
- Excel.QueryTable.WebTables
ms.assetid: d60eb457-6276-2d86-bbd8-c2050b0695c9
ms.date: 05/03/2019
localization_priority: Normal
---


# QueryTable.WebTables property (Excel)

Returns or sets a comma-delimited list of table names or table index numbers when you import a webpage into a query table. Read/write **String**.


## Syntax

_expression_.**WebTables**

_expression_ A variable that represents a **[QueryTable](Excel.QueryTable.md)** object.


## Remarks

Use this property only when the query table's **[QueryType](Excel.QueryTable.QueryType.md)** property is set to **xlWebQuery**, the query returns an HTML document, and the value of the **[WebSelectionType](Excel.QueryTable.WebSelectionType.md)** property is **xlSpecifiedTables**.

If you import data by using the user interface, data from a web query or a text query is imported as a **QueryTable** object, while all other external data is imported as a **[ListObject](Excel.ListObject.md)** object.

If you import data by using the object model, data from a web query or a text query must be imported as a **QueryTable**, while all other external data can be imported as either a **ListObject** or a **QueryTable**.

The **WebTables** property applies only to **QueryTable** objects.


## Example

This example adds a new web query table to the first worksheet in the first workbook, and then imports data from the first and second tables on the webpage.

```vb
Set shFirstQtr = Workbooks(1).Worksheets(1) 
Set qtQtrResults = shFirstQtr.QueryTables _ 
 .Add(Connection := "URL;https://datasvr/98q1/19980331.htm", _ 
 Destination := shFirstQtr.Cells(1,1)) 
With qtQtrResults 
 .WebFormatting = xlNone 
 .WebSelectionType = xlSpecifiedTables 
 .WebTables = "1,2" 
 .Refresh 
End With 

```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
