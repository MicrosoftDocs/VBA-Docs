---
title: Application.ReferenceStyle property (Excel)
keywords: vbaxl10.chm133197
f1_keywords:
- vbaxl10.chm133197
ms.prod: excel
api_name:
- Excel.Application.ReferenceStyle
ms.assetid: 86c4931b-ab1a-0363-d048-5195707a952b
ms.date: 04/05/2019
localization_priority: Normal
---


# Application.ReferenceStyle property (Excel)

Returns or sets how Microsoft Excel displays cell references and row and column headings in either A1 or R1C1 reference style. Read/write **[XlReferenceStyle](Excel.XlReferenceStyle.md)**.


## Syntax

_expression_.**ReferenceStyle**

_expression_ A variable that represents an **[Application](Excel.Application(object).md)** object.


## Remarks

**XlReferenceStyle** can be one of these constants:

- **xlA1**
- **xlR1C1**

## Example

This example displays the current reference style.

```vb
If Application.ReferenceStyle = xlR1C1 Then 
 MsgBox ("Microsoft Excel is using R1C1 references") 
Else 
 MsgBox ("Microsoft Excel is using A1 references") 
End If
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]