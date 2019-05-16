---
title: TextFrame.AutoMargins property (Excel)
keywords: vbaxl10.chm644083
f1_keywords:
- vbaxl10.chm644083
ms.prod: excel
api_name:
- Excel.TextFrame.AutoMargins
ms.assetid: a91ecac5-c907-8ae1-a0b8-1569f2466adf
ms.date: 05/17/2019
localization_priority: Normal
---


# TextFrame.AutoMargins property (Excel)

Returns or sets whether Excel automatically calculates text frame margins. Read/write.


## Syntax

_expression_.**AutoMargins**

_expression_ A variable that represents a **[TextFrame](Excel.TextFrame.md)** object.


## Return value

**Boolean**


## Remarks

**True** if Excel automatically calculates text frame margins; otherwise, **False**. 

When this property is **True**, the **[MarginLeft](Excel.TextFrame.MarginLeft.md)**, **[MarginRight](Excel.TextFrame.MarginRight.md)**, **[MarginTop](Excel.TextFrame.MarginTop.md)**, and **[MarginBottom](Excel.TextFrame.MarginBottom.md)** properties are ignored.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]