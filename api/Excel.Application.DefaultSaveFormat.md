---
title: Application.DefaultSaveFormat property (Excel)
keywords: vbaxl10.chm133217
f1_keywords:
- vbaxl10.chm133217
ms.prod: excel
api_name:
- Excel.Application.DefaultSaveFormat
ms.assetid: bb5c50db-8ba3-f79a-4577-f293ebc52b50
ms.date: 04/04/2019
localization_priority: Normal
---


# Application.DefaultSaveFormat property (Excel)

Returns or sets the default format for saving files. For a list of valid constants, see the **[FileFormat](Excel.Workbook.FileFormat.md)** property. Read/write **Long**.


## Syntax

_expression_.**DefaultSaveFormat**

_expression_ A variable that represents an **[Application](Excel.Application(object).md)** object.


## Example

This example sets the default format for saving files.


```vb
Application.DefaultSaveFormat = xlExcel4Workbook
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]