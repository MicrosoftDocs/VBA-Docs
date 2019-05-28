---
title: Workbook.EnvelopeVisible property (Excel)
keywords: vbaxl10.chm199191
f1_keywords:
- vbaxl10.chm199191
ms.prod: excel
api_name:
- Excel.Workbook.EnvelopeVisible
ms.assetid: d511a75a-ddd1-64f5-a09b-720657f64c09
ms.date: 05/29/2019
localization_priority: Normal
---


# Workbook.EnvelopeVisible property (Excel)

**True** if the email composition header and the envelope toolbar are both visible. Read/write **Boolean**.


## Syntax

_expression_.**EnvelopeVisible**

_expression_ A variable that represents a **[Workbook](Excel.Workbook.md)** object.


## Example

This example checks to see whether the email composition header and the envelope toolbar are visible in the first workbook. If they are visible, the example then sets the variable `strSubject` to the text of the email subject line.

```vb
If Workbooks(1).EnvelopeVisible = True Then 
 strSubject = "Please read: Review immediately" 
End If
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]