---
title: QueryTable.TextFileSemicolonDelimiter property (Excel)
keywords: vbaxl10.chm518104
f1_keywords:
- vbaxl10.chm518104
ms.prod: excel
api_name:
- Excel.QueryTable.TextFileSemicolonDelimiter
ms.assetid: 61a4ea08-aadd-6cf5-b810-448fe00b68f5
ms.date: 05/03/2019
localization_priority: Normal
---


# QueryTable.TextFileSemicolonDelimiter property (Excel)

**True** if the semicolon is the delimiter when you import a text file into a query table, and if the value of the **[TextFileParseType](Excel.QueryTable.TextFileParseType.md)** property is **xlDelimited**. The default value is **False**. Read/write **Boolean**.


## Syntax

_expression_.**TextFileSemicolonDelimiter**

_expression_ A variable that represents a **[QueryTable](Excel.QueryTable.md)** object.


## Remarks

Use this property only when your query table is based on data from a text file (with the **[QueryType](Excel.QueryTable.QueryType.md)** property set to **xlTextImport**).

If you import data by using the user interface, data from a web query or a text query is imported as a **QueryTable** object, while all other external data is imported as a **[ListObject](Excel.ListObject.md)** object.

If you import data by using the object model, data from a web query or a text query must be imported as a **QueryTable**, while all other external data can be imported as either a **ListObject** or a **QueryTable**.

The **TextFileSemicolonDelimiter** property applies only to **QueryTable** objects.


## Example

This example sets the semicolon to be the delimiter in the query table on the first worksheet in the first workbook, and then refreshes the query table.

```vb
Set shFirstQtr = Workbooks(1).Worksheets(1) 
Set qtQtrResults = shFirstQtr.QueryTables _ 
 .Add(Connection := "TEXT;C:\My Documents\19980331.txt", _ 
 Destination := shFirstQtr.Cells(1,1)) 
With qtQtrResults 
 .TextFileParseType = xlDelimited 
 .TextFileSemicolonDelimiter = True 
 .Refresh 
End With
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]