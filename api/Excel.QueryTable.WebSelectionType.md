---
title: QueryTable.WebSelectionType property (Excel)
keywords: vbaxl10.chm518122
f1_keywords:
- vbaxl10.chm518122
ms.prod: excel
api_name:
- Excel.QueryTable.WebSelectionType
ms.assetid: f0068811-96f8-55c6-a18d-29af4ae3a0e2
ms.date: 06/08/2017
localization_priority: Normal
---


# QueryTable.WebSelectionType property (Excel)

Returns or sets a value that determines whether an entire Web page, all tables on the Web page, or only specific tables on the Web page are imported into a query table. Read/write  **[xlWebSelectionType](Excel.XlWebSelectionType.md)**.


## Syntax

_expression_. `WebSelectionType`

_expression_ A variable that represents a [QueryTable](Excel.QueryTable.md) object.


## Remarks

Use this property only when the query table's  **[QueryType](Excel.QueryTable.QueryType.md)** property is set to **xlWebQuery** and the query returns an HTML document.

If the value of this property is  **xlSpecifiedTables** , you can use the **[WebTables](Excel.QueryTable.WebTables.md)** property to specify the tables to be imported.



|XlWebSelectionType can be one of these XlWebSelectionType constants.|
| **xlEntirePage**|
| **xlAllTables**_default_|
| **xlSpecifiedTables**|

If you import data using the user interface, data from a Web query or a text query is imported as a  **[QueryTable](Excel.QueryTable.md)** object, while all other external data is imported as a **[ListObject](Excel.ListObject.md)** object.

If you import data using the object model, data from a Web query or a text query must be imported as a  **QueryTable** , while all other external data can be imported as either a **ListObject** or a **QueryTable**.

The  **WebSelectionType** property applies only to **QueryTable** objects.


## Example

This example adds a new Web query table to the first worksheet in the first workbook and then imports data from the first and second tables in the Web page.


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


## See also


[QueryTable Object](Excel.QueryTable.md)

