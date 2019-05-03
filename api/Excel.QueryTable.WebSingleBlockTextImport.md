---
title: QueryTable.WebSingleBlockTextImport property (Excel)
keywords: vbaxl10.chm518126
f1_keywords:
- vbaxl10.chm518126
ms.prod: excel
api_name:
- Excel.QueryTable.WebSingleBlockTextImport
ms.assetid: 044de013-a065-86a3-b910-d4dec0a761b8
ms.date: 05/03/2019
localization_priority: Normal
---


# QueryTable.WebSingleBlockTextImport property (Excel)

**True** if data from the HTML `<PRE>` tags on the specified webpage is processed all at once when you import the page into a query table. **False** if the data is imported in blocks of contiguous rows so that header rows will be recognized as such. The default value is **False**. Read/write **Boolean**.


## Syntax

_expression_.**WebSingleBlockTextImport**

_expression_ A variable that represents a **[QueryTable](Excel.QueryTable.md)** object.


## Remarks

Use this property only when the query table's **[QueryType](Excel.QueryTable.QueryType.md)** property is set to **xlWebQuery**, and the query returns an HTML document.

If you import data by using the user interface, data from a web query or a text query is imported as a **QueryTable** object, while all other external data is imported as a **[ListObject](Excel.ListObject.md)** object.

If you import data by using the object model, data from a web query or a text query must be imported as a **QueryTable**, while all other external data can be imported as either a **ListObject** or a **QueryTable**.

The **WebSingleBlockTextImport** property applies only to **QueryTable** objects.


## Example

This example adds a new web query table to the first worksheet in the first workbook, and then imports all the HTML `<PRE>` tag data all at once.

```vb
Set shFirstQtr = Workbooks(1).Worksheets(1) 
Set qtQtrResults = shFirstQtr.QueryTables _ 
 .Add(Connection := "URL;https://datasvr/98q1/19980331.htm", _ 
 Destination := shFirstQtr.Cells(1,1)) 
With qtQtrResults 
 .WebSingleBlockTextImport = True 
 .Refresh 
End With
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]