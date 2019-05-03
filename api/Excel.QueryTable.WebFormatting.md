---
title: QueryTable.WebFormatting property (Excel)
keywords: vbaxl10.chm518123
f1_keywords:
- vbaxl10.chm518123
ms.prod: excel
api_name:
- Excel.QueryTable.WebFormatting
ms.assetid: 3ba96959-1c50-8cc0-0025-b5006b1ad62c
ms.date: 05/03/2019
localization_priority: Normal
---


# QueryTable.WebFormatting property (Excel)

Returns or sets a value that determines how much formatting from a webpage, if any, is applied when you import the page into a query table. Read/write **[XlWebFormatting](Excel.XlWebFormatting.md)**.


## Syntax

_expression_.**WebFormatting**

_expression_ A variable that represents a **[QueryTable](Excel.QueryTable.md)** object.


## Remarks

Use this property only when the query table's **[QueryType](Excel.QueryTable.QueryType.md)** property is set to **xlWebQuery**, and the query returns an HTML document.

If you import data by using the user interface, data from a web query or a text query is imported as a **QueryTable** object, while all other external data is imported as a **[ListObject](Excel.ListObject.md)** object.

If you import data by using the object model, data from a web query or a text query must be imported as a **QueryTable**, while all other external data can be imported as either a **ListObject** or a **QueryTable**.

The **WebFormatting** property applies only to **QueryTable** objects.


## Example

This example adds a new web query table to the first worksheet in the first workbook, imports all the webpage formatting applied to the data, and then refreshes the query table.

```vb
Set shFirstQtr = Workbooks(1).Worksheets(1) 
Set qtQtrResults = shFirstQtr.QueryTables _ 
 .Add(Connection := "URL;https://datasvr/98q1/19980331.htm", _ 
 Destination := shFirstQtr.Cells(1,1)) 
With qtQtrResults 
 .WebFormatting = xlAll 
 .Refresh 
End With
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]