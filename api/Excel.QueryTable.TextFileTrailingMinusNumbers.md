---
title: QueryTable.TextFileTrailingMinusNumbers property (Excel)
keywords: vbaxl10.chm518134
f1_keywords:
- vbaxl10.chm518134
ms.prod: excel
api_name:
- Excel.QueryTable.TextFileTrailingMinusNumbers
ms.assetid: 4e2257b2-fc88-145b-d307-35b6877d390b
ms.date: 06/08/2017
localization_priority: Normal
---


# QueryTable.TextFileTrailingMinusNumbers property (Excel)

 **True** for Microsoft Excel to treat numbers imported as text that begin with a "-" symbol as a negative symbol. **False** for Excel to treat numbers imported as text that begin with a "-" symbol as text. Read/write **Boolean**.


## Syntax

_expression_. `TextFileTrailingMinusNumbers`

_expression_ A variable that represents a [QueryTable](Excel.QueryTable.md) object.


## Remarks

If you import data using the user interface, data from a web query or a text query is imported as a  **[QueryTable](Excel.QueryTable.md)** object, while all other external data is imported as a **[ListObject](Excel.ListObject.md)** object.

If you import data using the object model, data from a web query or a text query must be imported as a  **QueryTable**, while all other external data can be imported as either a **ListObject** or a **QueryTable**.

The  **TextFileTrailingMinusNumbers** property applies only to **QueryTable** objects.


## Example

In this example, Microsoft Excel determines the setting for cell A1, treating numbers imported as text that begin with a "-" symbol. This example assumes a  **[QueryTable](Excel.QueryTable.md)** object exists on the active worksheet.


```vb
Sub CheckQueryTableSetting() 
 
 ' Determine setting for TextFileTrailingMinusNumbers 
 If Range("A1").QueryTable.TextFileTrailingMinusNumbers = True Then 
 MsgBox "Numbers imported as text that begin with a '-' symbol " & _ 
 "will be treated as a negative symbol." 
 Else 
 MsgBox "Numbers imported as text that begin with a '-' symbol " & _ 
 "will not be treated as a negative symbol." 
 End If 
 
End Sub
```


## See also


[QueryTable Object](Excel.QueryTable.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]