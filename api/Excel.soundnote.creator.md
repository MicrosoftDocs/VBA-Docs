---
title: SoundNote.Creator property (Excel)
keywords: vbaxl10.chm257074
f1_keywords:
- vbaxl10.chm257074
ms.assetid: 3b59f17a-56ca-16b0-4094-ead8e42ffd89
ms.date: 05/16/2019
ms.localizationpriority: medium
---


# SoundNote.Creator property (Excel)

> [!NOTE] 
> This property should not be used. Sound notes have been removed from Microsoft Excel.

Returns a 32-bit integer that indicates the application in which this object was created. Read-only **Long**.


## Syntax

_expression_.**Creator**

_expression_ A variable that represents a **SoundNote** object.


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]