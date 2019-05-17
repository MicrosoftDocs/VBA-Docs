---
title: Windows.Creator property (Excel)
keywords: vbaxl10.chm353074
f1_keywords:
- vbaxl10.chm353074
ms.prod: excel
api_name:
- Excel.Windows.Creator
ms.assetid: f27724b1-4ce1-1b90-9aa3-704e491575f7
ms.date: 05/18/2019
localization_priority: Normal
---


# Windows.Creator property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only **Long**.


## Syntax

_expression_.**Creator**

_expression_ A variable that represents a **[Windows](Excel.Windows.md)** object.


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]