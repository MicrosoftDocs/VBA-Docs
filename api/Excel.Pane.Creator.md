---
title: Pane.Creator property (Excel)
keywords: vbaxl10.chm359074
f1_keywords:
- vbaxl10.chm359074
ms.prod: excel
api_name:
- Excel.Pane.Creator
ms.assetid: 600f4ae4-8fb6-c81e-c29c-654f72235bc1
ms.date: 05/03/2019
localization_priority: Normal
---


# Pane.Creator property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only **Long**.


## Syntax

_expression_.**Creator**

_expression_ A variable that represents a **[Pane](Excel.Pane.md)** object.


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]