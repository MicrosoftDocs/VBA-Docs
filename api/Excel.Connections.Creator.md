---
title: Connections.Creator property (Excel)
keywords: vbaxl10.chm775074
f1_keywords:
- vbaxl10.chm775074
api_name:
- Excel.Connections.Creator
ms.assetid: eb334a7c-d286-c1a0-c4d3-a4a2fe5be7c2
ms.date: 04/23/2019
ms.localizationpriority: medium
---


# Connections.Creator property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only **Long**.


## Syntax

_expression_.**Creator**

_expression_ A variable that represents a **[Connections](Excel.Connections.md)** object.


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]