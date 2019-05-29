---
title: Worksheet.EnableAutoFilter property (Excel)
keywords: vbaxl10.chm175094
f1_keywords:
- vbaxl10.chm175094
ms.prod: excel
api_name:
- Excel.Worksheet.EnableAutoFilter
ms.assetid: bff7829a-30f7-3248-e694-ac48621aed31
ms.date: 05/30/2019
localization_priority: Normal
---


# Worksheet.EnableAutoFilter property (Excel)

**True** if AutoFilter arrows are enabled when user-interface-only protection is turned on. Read/write **Boolean**.


## Syntax

_expression_.**EnableAutoFilter**

_expression_ A variable that represents a **[Worksheet](Excel.Worksheet.md)** object.


## Example

This example enables the AutoFilter arrows on a protected worksheet.

```vb
ActiveSheet.EnableAutoFilter = True 
ActiveSheet.Protect contents:=True, userInterfaceOnly:=True
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]