---
title: Ranges.Creator property (Excel)
keywords: vbaxl10.chm827074
f1_keywords:
- vbaxl10.chm827074
ms.prod: excel
api_name:
- Excel.Ranges.Creator
ms.assetid: a44c830a-4164-6fa5-b316-399019ed77ab
ms.date: 05/11/2019
localization_priority: Normal
---


# Ranges.Creator property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only **Long**.


## Syntax

_expression_.**Creator**

_expression_ A variable that represents a **[Ranges](Excel.Ranges.md)** object.


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]