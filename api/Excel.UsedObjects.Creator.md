---
title: UsedObjects.Creator property (Excel)
keywords: vbaxl10.chm677074
f1_keywords:
- vbaxl10.chm677074
ms.prod: excel
api_name:
- Excel.UsedObjects.Creator
ms.assetid: 18526467-0a48-b372-bb34-6a810da25ce2
ms.date: 05/18/2019
localization_priority: Normal
---


# UsedObjects.Creator property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only **Long**.


## Syntax

_expression_.**Creator**

_expression_ A variable that represents a **[UsedObjects](Excel.UsedObjects.md)** object.


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]