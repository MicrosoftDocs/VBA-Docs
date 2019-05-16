---
title: TableStyles.Creator property (Excel)
keywords: vbaxl10.chm840074
f1_keywords:
- vbaxl10.chm840074
ms.prod: excel
api_name:
- Excel.TableStyles.Creator
ms.assetid: dc129be8-a2d9-b194-8759-88298470afe7
ms.date: 05/17/2019
localization_priority: Normal
---


# TableStyles.Creator property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only **Long**.


## Syntax

_expression_.**Creator**

_expression_ A variable that represents a **[TableStyles](Excel.TableStyles.md)** object.


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]