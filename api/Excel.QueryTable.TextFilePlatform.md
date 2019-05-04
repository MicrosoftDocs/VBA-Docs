---
title: QueryTable.TextFilePlatform property (Excel)
keywords: vbaxl10.chm518098
f1_keywords:
- vbaxl10.chm518098
ms.prod: excel
api_name:
- Excel.QueryTable.TextFilePlatform
ms.assetid: 2fb3dbb5-919e-2e27-9fbf-8feaa107c2a7
ms.date: 05/03/2019
localization_priority: Normal
---


# QueryTable.TextFilePlatform property (Excel)

Returns or sets the origin of the text file that you are importing into the query table. This property determines which code page is used during the data import. Read/write **[XlPlatform](Excel.XlPlatform.md)**.


## Syntax

_expression_.**TextFilePlatform**

_expression_ A variable that represents a **[QueryTable](Excel.QueryTable.md)** object.


## Remarks

The default value is the current setting of the **File Origin** option in the **Text File Import Wizard**.

Use this property only when your query table is based on data from a text file (with the **[QueryType](Excel.QueryTable.QueryType.md)** property set to **xlTextImport**).

If you import data by using the user interface, data from a web query or a text query is imported as a **QueryTable** object, while all other external data is imported as a **[ListObject](Excel.ListObject.md)** object.

If you import data by using the object model, data from a web query or a text query must be imported as a **QueryTable**, while all other external data can be imported as either a **ListObject** or a **QueryTable**.

The **TextFilePlatform** property applies only to **QueryTable** objects.


## Example

This example imports an MS-DOS text file into the query table on the first worksheet in the first workbook, and then it refreshes the query table.

```vb
Set shFirstQtr = Workbooks(1).Worksheets(1) 
Set qtQtrResults = shFirstQtr.QueryTables _ 
 .Add(Connection := "TEXT;C:\My Documents\19980331.txt", _ 
 Destination := shFirstQtr.Cells(1,1)) 
With qtQtrResults 
 .TextFilePlatform = xlMSDOS 
 .TextFileParseType = xlDelimited 
 .TextFileTabDelimiter = True 
 .Refresh 
End With
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
