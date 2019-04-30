---
title: ModelFormatGeneral.Creator property (Excel)
keywords: vbaxl10.chm981074
f1_keywords:
- vbaxl10.chm981074
ms.assetid: 828ced24-d35d-bee5-c9a6-b63e102c8cfb
ms.date: 05/01/2019
ms.prod: excel
localization_priority: Normal
---


# ModelFormatGeneral.Creator property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only **Long**.


## Syntax

_expression_.**Creator**

_expression_ A variable that represents a **[ModelFormatGeneral](Excel.modelformatgeneral.md)** object.


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]