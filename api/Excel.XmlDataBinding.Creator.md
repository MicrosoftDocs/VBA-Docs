---
title: XmlDataBinding.Creator property (Excel)
keywords: vbaxl10.chm747074
f1_keywords:
- vbaxl10.chm747074
ms.prod: excel
api_name:
- Excel.XmlDataBinding.Creator
ms.assetid: 1d03c514-abed-3987-0f5a-652f5befe972
ms.date: 05/21/2019
localization_priority: Normal
---


# XmlDataBinding.Creator property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only **Long**.


## Syntax

_expression_.**Creator**

_expression_ A variable that represents an **[XmlDataBinding](Excel.XmlDataBinding.md)** object.


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]