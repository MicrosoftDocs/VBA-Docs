---
title: Application.Creator property (Excel)
keywords: vbaxl10.chm182074
f1_keywords:
- vbaxl10.chm182074
api_name:
- Excel.Application.Creator
ms.assetid: 92ceed4a-4e47-18d5-6023-f1018eefd071
ms.date: 04/04/2019
ms.localizationpriority: medium
---


# Application.Creator property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only **Long**.


## Syntax

_expression_.**Creator**

_expression_ An expression that returns an **[Application](Excel.Application(object).md)** object.


## Return value

XlCreator


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]