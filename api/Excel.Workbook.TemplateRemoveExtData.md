---
title: Workbook.TemplateRemoveExtData property (Excel)
keywords: vbaxl10.chm199171
f1_keywords:
- vbaxl10.chm199171
ms.prod: excel
api_name:
- Excel.Workbook.TemplateRemoveExtData
ms.assetid: 9851df1d-4e83-525a-8a43-bd84b0a94c74
ms.date: 05/29/2019
localization_priority: Normal
---


# Workbook.TemplateRemoveExtData property (Excel)

**True** if external data references are removed when the workbook is saved as a template. Read/write **Boolean**.


## Syntax

_expression_.**TemplateRemoveExtData**

_expression_ A variable that represents a **[Workbook](Excel.Workbook.md)** object.


## Example

This example saves the workbook as a template that contains no external data.

```vb
With ThisWorkbook 
 .TemplateRemoveExtData = True 
 .SaveAs "current", xlTemplate 
 .TemplateRemoveExtData = False 
End With
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]