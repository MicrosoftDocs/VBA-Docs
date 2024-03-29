---
title: CubeField.Creator property (Excel)
keywords: vbaxl10.chm667074
f1_keywords:
- vbaxl10.chm667074
api_name:
- Excel.CubeField.Creator
ms.assetid: 2534f870-90cd-e3ab-b1fd-d63455a75809
ms.date: 04/23/2019
ms.localizationpriority: medium
---


# CubeField.Creator property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only **Long**.


## Syntax

_expression_.**Creator**

_expression_ A variable that represents a **[CubeField](Excel.CubeField.md)** object.


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]