---
title: OLEObject.Creator property (Excel)
keywords: vbaxl10.chm414074
f1_keywords:
- vbaxl10.chm414074
api_name:
- Excel.OLEObject.Creator
ms.assetid: 6bbbaad2-30f5-c443-c6ab-b6c375a7810f
ms.date: 05/02/2019
ms.localizationpriority: medium
---


# OLEObject.Creator property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only **Long**.


## Syntax

_expression_.**Creator**

_expression_ A variable that represents an **[OLEObject](Excel.OLEObject.md)** object.


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]