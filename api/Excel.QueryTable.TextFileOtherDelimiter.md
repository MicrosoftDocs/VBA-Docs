---
title: QueryTable.TextFileOtherDelimiter property (Excel)
keywords: vbaxl10.chm518107
f1_keywords:
- vbaxl10.chm518107
ms.prod: excel
api_name:
- Excel.QueryTable.TextFileOtherDelimiter
ms.assetid: e632984a-4316-4e65-754f-01a2c77d5cad
ms.date: 05/03/2019
localization_priority: Normal
---


# QueryTable.TextFileOtherDelimiter property (Excel)

Returns or sets the character used as the delimiter when you import a text file into a query table. The default value is **null**. Read/write **String**.


## Syntax

_expression_.**TextFileOtherDelimiter**

_expression_ A variable that represents a **[QueryTable](Excel.QueryTable.md)** object.


## Remarks

Use this property only when your query table is based on data from a text file (with the **[QueryType](Excel.QueryTable.QueryType.md)** property set to **xlTextImport**), and only if the value of the **[TextFileParseType](Excel.QueryTable.TextFileParseType.md)** property is **xlDelimited**.

If you specify more than one character in the string, only the first character is used.

If you import data by using the user interface, data from a web query or a text query is imported as a **QueryTable** object, while all other external data is imported as a **[ListObject](Excel.ListObject.md)** object.

If you import data by using the object model, data from a web query or a text query must be imported as a **QueryTable**, while all other external data can be imported as either a **ListObject** or a **QueryTable**.

The **TextFileOtherDelimiter** property applies only to **QueryTable** objects.


## Example

This example sets the pound character (#) to be the delimiter for the query table on the first worksheet in the first workbook, and then it refreshes the query table.

```vb
Set shFirstQtr = Workbooks(1).Worksheets(1) 
Set qtQtrResults = shFirstQtr.QueryTables _ 
 .Add(Connection := "TEXT;C:\My Documents\19980331.txt", _ 
 Destination := shFirstQtr.Cells(1,1)) 
With qtQtrResults 
 .TextFileParseType = xlDelimited 
 .TextFileOtherDelimiter = "#" 
 .Refresh 
End With
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]