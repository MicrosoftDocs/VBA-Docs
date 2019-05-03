---
title: QueryTable.TextFilePromptOnRefresh property (Excel)
keywords: vbaxl10.chm518115
f1_keywords:
- vbaxl10.chm518115
ms.prod: excel
api_name:
- Excel.QueryTable.TextFilePromptOnRefresh
ms.assetid: 3fe619b9-2bc8-46f4-4e18-655e9cf5a61f
ms.date: 05/03/2019
localization_priority: Normal
---


# QueryTable.TextFilePromptOnRefresh property (Excel)

**True** if you want to specify the name of the imported text file each time the query table is refreshed. The **Import Text File** dialog box allows you to specify the path and file name. The default value is **False**. Read/write **Boolean**.


## Syntax

_expression_.**TextFilePromptOnRefresh**

_expression_ A variable that represents a **[QueryTable](Excel.QueryTable.md)** object.


## Remarks

Use this property only when your query table is based on data from a text file (with the **[QueryType](Excel.QueryTable.QueryType.md)** property set to **xlTextImport**).

If the value of this property is **True**, the dialog box doesn't appear the first time that a query table is refreshed.

The default value is **True** in the user interface.

If you import data by using the user interface, data from a web query or a text query is imported as a **QueryTable** object, while all other external data is imported as a **[ListObject](Excel.ListObject.md)** object.

If you import data by using the object model, data from a web query or a text query must be imported as a **QueryTable**, while all other external data can be imported as either a **ListObject** or a **QueryTable**.

The **TextFilePromptOnRefresh** property applies only to **QueryTable** objects.


## Example

This example prompts the user for the name of the text file whenever the query table on the first worksheet in the first workbook is refreshed.

```vb
Set shFirstQtr = Workbooks(1).Worksheets(1) 
Set qtQtrResults = shFirstQtr.QueryTables _ 
 .Add(Connection := "TEXT;C:\My Documents\19980331.txt", _ 
 Destination := shFirstQtr.Cells(1,1)) 
With qtQtrResults 
 .TextFileParseType = xlDelimited 
 .TextFilePromptOnRefresh = True 
 .TextFileTabDelimiter = True 
 .Refresh 
End With
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]