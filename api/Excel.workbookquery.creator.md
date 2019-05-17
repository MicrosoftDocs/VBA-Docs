---
title: WorkbookQuery.Creator property (Excel)
keywords: vbaxl10.chm973074
f1_keywords:
- vbaxl10.chm973074
ms.assetid: 82e257ca-9e3f-0acc-66a7-84f7e7e07ff8
ms.date: 05/18/2019
ms.prod: excel
localization_priority: Normal
---


# WorkbookQuery.Creator property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only **Long**.


## Syntax

_expression_.**Creator**

_expression_ A variable that represents a **[WorkbookQuery](Excel.WorkbookQuery.md)** object.


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]