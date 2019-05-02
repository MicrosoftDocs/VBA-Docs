---
title: QueryTable.Creator property (Excel)
keywords: vbaxl10.chm517074
f1_keywords:
- vbaxl10.chm517074
ms.prod: excel
api_name:
- Excel.QueryTable.Creator
ms.assetid: 6384b8d4-295c-1566-9405-a7450551b4f1
ms.date: 05/03/2019
localization_priority: Normal
---


# QueryTable.Creator property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only **Long**.


## Syntax

_expression_.**Creator**

_expression_ A variable that represents a **[QueryTable](Excel.QueryTable.md)** object.


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.

Data from web queries or text queries is imported as a **QueryTable** object, while all other external data is imported as a **[ListObject](Excel.ListObject.md)** object. You can use the **[QueryTable](Excel.ListObject.QueryTable.md)** property of the **ListObject** to access the **Creator** property.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]