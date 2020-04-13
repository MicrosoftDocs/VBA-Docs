---
title: Row.GetValues method (Outlook)
keywords: vbaol11.chm2244
f1_keywords:
- vbaol11.chm2244
ms.prod: outlook
api_name:
- Outlook.Row.GetValues
ms.assetid: 1f92e0ab-9ba8-9cc6-51e8-05cc145a93bf
ms.date: 06/08/2017
localization_priority: Normal
---


# Row.GetValues method (Outlook)

Obtains a one-dimensional array containing the values for all columns at the  **[Row](Outlook.Row.md)** in the parent **[Table](Outlook.Table.md)**.


## Syntax

_expression_. `GetValues`

_expression_ A variable that represents a [Row](Outlook.Row.md) object.


## Return value

A **Variant** that represents an array of values for all the columns at that **Row** in the **Table**.


## Remarks

 **GetValues** is a helper method that facilitates fetching all the column values in the **Row** in a single call.

Since the array is zero-based, the length of the array is the number of columns in the  **Row** minus one.

Values returned in the array are of the same type as the values in the parent  **Table**. This means that binary properties in the **Table** are returned as arrays of bytes. For date-time properties, if a **[Column](Outlook.Column.md)** is a default column or if it has been added using an explicit built-in property name, then its value in the **Table** and in the array are expressed in local time. If the **Column** has been added to the **Table** using a namespace reference, then its value in the **Table** and in the array are expressed in Coordinated Universal Time (UTC). For more information on referencing properties by namespace, see [Referencing Properties by Namespace](../outlook/How-to/Navigation/referencing-properties-by-namespace.md). 


## See also


[Row Object](Outlook.Row.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]