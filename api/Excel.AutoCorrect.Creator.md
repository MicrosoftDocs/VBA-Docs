---
title: AutoCorrect.Creator property (Excel)
keywords: vbaxl10.chm544074
f1_keywords:
- vbaxl10.chm544074
ms.prod: excel
api_name:
- Excel.AutoCorrect.Creator
ms.assetid: 25c3b228-cfac-8703-acd9-533cf86387cb
ms.date: 04/06/2019
localization_priority: Normal
---


# AutoCorrect.Creator property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only **Long**.


## Syntax

_expression_.**Creator**

_expression_ A variable that represents an **[AutoCorrect](Excel.AutoCorrect(object).md)** object.


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]