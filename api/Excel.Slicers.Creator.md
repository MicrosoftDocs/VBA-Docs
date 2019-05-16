---
title: Slicers.Creator property (Excel)
keywords: vbaxl10.chm902074
f1_keywords:
- vbaxl10.chm902074
ms.prod: excel
api_name:
- Excel.Slicers.Creator
ms.assetid: c02ffedb-014b-1fd2-544c-5cd6dc83912a
ms.date: 05/16/2019
localization_priority: Normal
---


# Slicers.Creator property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only **Long**.


## Syntax

_expression_.**Creator**

_expression_ A variable that represents a **[Slicers](Excel.Slicers.md)** object.


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]