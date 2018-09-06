---
title: QueryTable.TextFileSpaceDelimiter Property (Excel)
keywords: vbaxl10.chm518106
f1_keywords:
- vbaxl10.chm518106
ms.prod: excel
api_name:
- Excel.QueryTable.TextFileSpaceDelimiter
ms.assetid: d2c5fb8a-f235-d6d4-a73c-29477ea24fe4
ms.date: 06/08/2017
---


# QueryTable.TextFileSpaceDelimiter Property (Excel)

 **True** if the space character is the delimiter when you import a text file into a query table. The default value is **False** . Read/write **Boolean** .


## Syntax

 _expression_. `TextFileSpaceDelimiter`

 _expression_ A variable that represents a [QueryTable](Excel.QueryTable.md) object.


## Remarks

Use this property only when your query table is based on data from a text file (with the  **[QueryType](Excel.QueryTable.QueryType.md)** property set to **xlTextImport**), and only if the value of the **[TextFileParseType](Excel.QueryTable.TextFileParseType.md)** property is **xlDelimited** .

If you import data using the user interface, data from a Web query or a text query is imported as a  **[QueryTable](Excel.QueryTable.md)** object, while all other external data is imported as a **[ListObject](Excel.ListObject.md)** object.

If you import data using the object model, data from a Web query or a text query must be imported as a  **QueryTable** , while all other external data can be imported as either a **ListObject** or a **QueryTable** .

The  **TextFileSpaceDelimiter** property applies only to **QueryTable** objects.


## Example

This example sets the space character to be the delimiter in the query table on the first worksheet in the first workbook and then refreshes the query table.


```vb
Set shFirstQtr = Workbooks(1).Worksheets(1) 
Set qtQtrResults = shFirstQtr.QueryTables _ 
 .Add(Connection := "TEXT;C:\My Documents\19980331.txt", _ 
 Destination := shFirstQtr.Cells(1,1)) 
With qtQtrResults 
 .TextFileParseType = xlDelimited 
 .TextFileSpaceDelimiter = True 
 .Refresh 
End With
```


## See also


[QueryTable Object](Excel.QueryTable.md)

