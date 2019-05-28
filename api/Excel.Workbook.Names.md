---
title: Workbook.Names property (Excel)
keywords: vbaxl10.chm199115
f1_keywords:
- vbaxl10.chm199115
ms.prod: excel
api_name:
- Excel.Workbook.Names
ms.assetid: 26be56ec-ea12-1600-602a-eb338d4a5a8b
ms.date: 05/29/2019
localization_priority: Normal
---


# Workbook.Names property (Excel)

Returns a **[Names](Excel.Names.md)** collection that represents all the names in the specified workbook (including all worksheet-specific names). Read-only **Names** object.


## Syntax

_expression_.**Names**

_expression_ A variable that represents a **[Workbook](Excel.Workbook.md)** object.


## Remarks

Using this property without an object qualifier is equivalent to using **ActiveWorkbook.Names**.


## Example

This example defines the name _myName_ for cell A1 on Sheet1.

```vb
ActiveWorkbook.Names.Add Name:="myName", RefersToR1C1:= _ 
 "=Sheet1!R1C1"
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
