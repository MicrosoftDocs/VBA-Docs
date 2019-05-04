---
title: QueryTable.WebConsecutiveDelimitersAsOne property (Excel)
keywords: vbaxl10.chm518128
f1_keywords:
- vbaxl10.chm518128
ms.prod: excel
api_name:
- Excel.QueryTable.WebConsecutiveDelimitersAsOne
ms.assetid: cc10dd93-2574-7575-3326-1d2992f4c731
ms.date: 05/03/2019
localization_priority: Normal
---


# QueryTable.WebConsecutiveDelimitersAsOne property (Excel)

**True** if consecutive delimiters are treated as a single delimiter when you import data from HTML `<PRE>` tags on a webpage into a query table, and if the data is to be parsed into columns. **False** if you want to treat consecutive delimiters as multiple delimiters. The default value is **True**. Read/write **Boolean**.


## Syntax

_expression_.**WebConsecutiveDelimitersAsOne**

_expression_ A variable that represents a **[QueryTable](Excel.QueryTable.md)** object.


## Remarks

Use this property only when the query table's **[QueryType](Excel.QueryTable.QueryType.md)** property is set to **xlWebQuery**, the query returns an HTML document, and the **[WebPreFormattedTextToColumns](Excel.QueryTable.WebPreFormattedTextToColumns.md)** property is set to **True**.

If you import data by using the user interface, data from a web query or a text query is imported as a **QueryTable** object, while all other external data is imported as a **[ListObject](Excel.ListObject.md)** object.

If you import data by using the object model, data from a web query or a text query must be imported as a **QueryTable**, while all other external data can be imported as either a **ListObject** or a **QueryTable**.

The **WebConsecutiveDelimitersAsOne** property applies only to **QueryTable** objects.


## Example

This example sets the space character to be the delimiter in the query table on the first worksheet in the first workbook, and then it refreshes the query table. Consecutive spaces are treated as a single space.

```vb
Set shFirstQtr = Workbooks(1).Worksheets(1) 
Set qtQtrResults = shFirstQtr.QueryTables _ 
 .Add(Connection := "URL;https://datasvr/98q1/19980331.htm", _ 
 Destination := shFirstQtr.Cells(1,1)) 
With qtQtrResults 
 .WebConsecutiveDelimitersAsOne = True 
 .Refresh 
End With
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]