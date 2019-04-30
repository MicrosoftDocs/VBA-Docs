---
title: LineFormat.Creator property (Excel)
ms.prod: excel
api_name:
- Excel.LineFormat.Creator
ms.assetid: afcb3c96-048f-e105-6c05-6bf455972284
ms.date: 04/30/2019
localization_priority: Normal
---


# LineFormat.Creator property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only **Long**.


## Syntax

_expression_.**Creator**

_expression_ A variable that represents a **[LineFormat](Excel.LineFormat.md)** object.


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]