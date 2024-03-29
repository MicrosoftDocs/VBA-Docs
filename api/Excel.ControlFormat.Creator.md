---
title: ControlFormat.Creator property (Excel)
keywords: vbaxl10.chm629074
f1_keywords:
- vbaxl10.chm629074
api_name:
- Excel.ControlFormat.Creator
ms.assetid: d3174b4f-70ad-4026-2205-8f71c8f1338a
ms.date: 04/23/2019
ms.localizationpriority: medium
---


# ControlFormat.Creator property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only **Long**.


## Syntax

_expression_.**Creator**

_expression_ A variable that represents a **[ControlFormat](Excel.ControlFormat.md)** object.


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]