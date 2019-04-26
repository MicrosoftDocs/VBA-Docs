---
title: Error.Creator property (Excel)
keywords: vbaxl10.chm701074
f1_keywords:
- vbaxl10.chm701074
ms.prod: excel
api_name:
- Excel.Error.Creator
ms.assetid: 88dd1cda-72a2-18bd-e6aa-83b5414767cd
ms.date: 04/26/2019
localization_priority: Normal
---


# Error.Creator property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only **Long**.


## Syntax

_expression_.**Creator**

_expression_ A variable that represents an **[Error](Excel.Error.md)** object.


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]