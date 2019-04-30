---
title: QueryTable.WebPreFormattedTextToColumns property (Excel)
keywords: vbaxl10.chm518125
f1_keywords:
- vbaxl10.chm518125
ms.prod: excel
api_name:
- Excel.QueryTable.WebPreFormattedTextToColumns
ms.assetid: 5365c5c8-9dc9-3140-c3cc-679bd0db4477
ms.date: 06/08/2017
localization_priority: Normal
---


# QueryTable.WebPreFormattedTextToColumns property (Excel)

Returns or sets whether data contained within HTML <PRE> tags on the webpage is parsed into columns when you import the page into a query table. The default is  **True**. Read/write **Boolean**.


## Syntax

_expression_. `WebPreFormattedTextToColumns`

_expression_ A variable that represents a **[QueryTable](Excel.QueryTable.md)** object.


## Remarks

This property is used only when the  **[QueryType](Excel.QueryTable.QueryType.md)** property of the query table is **xlWebQuery** and the query returns a HTML document.

If you import data using the user interface, data from a web query or a text query is imported as a  **[QueryTable](Excel.QueryTable.md)** object, while all other external data is imported as a **[ListObject](Excel.ListObject.md)** object.

If you import data using the object model, data from a web query or a text query must be imported as a  **QueryTable**, while all other external data can be imported as either a **ListObject** or a **QueryTable**.

The  **WebPreFormattedTextToColumns** property applies only to **QueryTable** objects.


## Example

This example adds a new Web query table to the first worksheet in the first workbook. Note that the example doesn't parse into columns any data located between the HTML <PRE> tags.


```vb
Set shFirstQtr = Workbooks(1).Worksheets(1) 
Set qtQtrResults = shFirstQtr.QueryTables _ 
 .Add(Connection := "URL;https://datasvr/98q1/19980331.htm", _ 
 Destination := shFirstQtr.Cells(1,1)) 
With qtQtrResults 
 .WebFormatting = xlNone 
 .WebPreFormattedTextToColumns = False 
 .Refresh 
End With
```


## See also


[QueryTable Object](Excel.QueryTable.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]