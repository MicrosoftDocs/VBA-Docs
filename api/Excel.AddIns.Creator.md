---
title: AddIns.Creator property (Excel)
keywords: vbaxl10.chm186074
f1_keywords:
- vbaxl10.chm186074
ms.prod: excel
api_name:
- Excel.AddIns.Creator
ms.assetid: 8fc7772e-1837-5336-9ae7-eca7f0dc14af
ms.date: 04/03/2019
localization_priority: Normal
---


# AddIns.Creator property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only **Long**.


## Syntax

_expression_.**Creator**

_expression_ An expression that returns an **[AddIns](Excel.AddIns.md)** object.


## Return value

XlCreator


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]