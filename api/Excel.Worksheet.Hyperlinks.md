---
title: Worksheet.Hyperlinks property (Excel)
keywords: vbaxl10.chm175140
f1_keywords:
- vbaxl10.chm175140
ms.prod: excel
api_name:
- Excel.Worksheet.Hyperlinks
ms.assetid: ac2fe50a-23a0-9982-d448-b18a91092624
ms.date: 05/30/2019
localization_priority: Normal
---


# Worksheet.Hyperlinks property (Excel)

Returns a **[Hyperlinks](Excel.Hyperlinks.md)** collection that represents the hyperlinks for the worksheet.


## Syntax

_expression_.**Hyperlinks**

_expression_ A variable that represents a **[Worksheet](Excel.Worksheet.md)** object.


## Example

This example checks to see whether any of the hyperlinks on worksheet one contain the word Microsoft.

```vb
For Each h in Worksheets(1).Hyperlinks 
 If Instr(h.Name, "Microsoft") <> 0 Then h.Follow 
Next
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]