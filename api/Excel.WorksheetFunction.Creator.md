---
title: WorksheetFunction.Creator property (Excel)
keywords: vbaxl10.chm136074
f1_keywords:
- vbaxl10.chm136074
ms.prod: excel
api_name:
- Excel.WorksheetFunction.Creator
ms.assetid: 142d1b93-b4cf-2d69-c2c3-48072e31032b
ms.date: 05/22/2019
localization_priority: Normal
---


# WorksheetFunction.Creator property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long**.


## Syntax

_expression_.**Creator**

_expression_ A variable that represents a **[WorksheetFunction](Excel.WorksheetFunction.md)** object.


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]