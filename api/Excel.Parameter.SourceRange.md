---
title: Parameter.SourceRange property (Excel)
keywords: vbaxl10.chm523077
f1_keywords:
- vbaxl10.chm523077
ms.prod: excel
api_name:
- Excel.Parameter.SourceRange
ms.assetid: 243ac075-24cc-549a-58fb-195d71dc6e68
ms.date: 05/03/2019
localization_priority: Normal
---


# Parameter.SourceRange property (Excel)

Returns a **[Range](Excel.Range(object).md)** object that represents the cell that contains the value of the specified query parameter. Read-only.


## Syntax

_expression_.**SourceRange**

_expression_ A variable that represents a **[Parameter](Excel.Parameter.md)** object.


## Example

This example changes the value of the cell used as the source range for the query.

```vb
Set qt = Sheets("sheet1").QueryTables(1) 
Set param1 = qt.Parameters(1) 
Set r = param1.SourceRange 
r.Value = "New York" 
qt.Refresh
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]