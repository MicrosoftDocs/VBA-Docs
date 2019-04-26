---
title: IconSetCondition.Creator property (Excel)
keywords: vbaxl10.chm811074
f1_keywords:
- vbaxl10.chm811074
ms.prod: excel
api_name:
- Excel.IconSetCondition.Creator
ms.assetid: 1d8441b4-b8df-9fe1-60f4-a3da1c9b2e57
ms.date: 04/27/2019
localization_priority: Normal
---


# IconSetCondition.Creator property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only **Long**.


## Syntax

_expression_.**Creator**

_expression_ A variable that represents an **[IconSetCondition](Excel.IconSetCondition.md)** object.


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]