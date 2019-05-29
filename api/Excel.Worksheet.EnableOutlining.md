---
title: Worksheet.EnableOutlining property (Excel)
keywords: vbaxl10.chm175096
f1_keywords:
- vbaxl10.chm175096
ms.prod: excel
api_name:
- Excel.Worksheet.EnableOutlining
ms.assetid: db849ddf-871d-19cd-9765-3194a8c1e34e
ms.date: 05/30/2019
localization_priority: Normal
---


# Worksheet.EnableOutlining property (Excel)

**True** if outlining symbols are enabled when user-interface-only protection is turned on. Read/write **Boolean**.


## Syntax

_expression_.**EnableOutlining**

_expression_ A variable that represents a **[Worksheet](Excel.Worksheet.md)** object.


## Example

This example enables outlining symbols on a protected worksheet.

```vb
ActiveSheet.EnableOutlining = True 
ActiveSheet.Protect contents:=True, userInterfaceOnly:=True
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]