---
title: DataLabels.Creator property (Excel)
keywords: vbaxl10.chm583074
f1_keywords:
- vbaxl10.chm583074
ms.prod: excel
api_name:
- Excel.DataLabels.Creator
ms.assetid: 86b4beaa-9a5f-ca3e-7e4f-78b905e94bae
ms.date: 04/23/2019
localization_priority: Normal
---


# DataLabels.Creator property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only **Long**.


## Syntax

_expression_.**Creator**

_expression_ A variable that represents a **[DataLabels](Excel.DataLabels(object).md)** object.


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]