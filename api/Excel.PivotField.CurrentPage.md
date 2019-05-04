---
title: PivotField.CurrentPage property (Excel)
keywords: vbaxl10.chm240077
f1_keywords:
- vbaxl10.chm240077
ms.prod: excel
api_name:
- Excel.PivotField.CurrentPage
ms.assetid: 4a59fe58-8f95-4cf3-d4a3-ab6ea6b24b8a
ms.date: 05/04/2019
localization_priority: Normal
---


# PivotField.CurrentPage property (Excel)

Returns or sets the current page showing for the page field (valid only for page fields). Read/write **[PivotItem](Excel.PivotItem.md)**.


## Syntax

_expression_.**CurrentPage**

_expression_ A variable that represents a **[PivotField](Excel.PivotField.md)** object.


## Example

This example returns the current page name for the PivotTable report on Sheet1 in the string variable `strPgName`.

```vb
Set pvtTable = Worksheets("Sheet1").Range("A3").PivotTable 
strPgName = pvtTable.PivotFields("Country").CurrentPage.Name
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
