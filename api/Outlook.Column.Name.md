---
title: Column.Name property (Outlook)
keywords: vbaol11.chm2749
f1_keywords:
- vbaol11.chm2749
ms.prod: outlook
api_name:
- Outlook.Column.Name
ms.assetid: e69a8a53-d348-2147-28cf-d41ea80bba61
ms.date: 06/08/2017
localization_priority: Normal
---


# Column.Name property (Outlook)

Returns a **String** value that represents the name of the **[Column](Outlook.Column.md)**. Read-only.


## Syntax

_expression_.**Name**

_expression_ A variable that represents a [Column](Outlook.Column.md) object.


## Remarks

The **Name** property is the default member of the **Column** object.

If the  **Column** is a default column in the **[Table](Outlook.Table.md)**, or if it has been added to the **Table** with the explicit built-in name for the property, the value of **Name** is the explicit built-in name (without any enclosing brackets) for the property. If the **Column** has been added to the **Table** with a property name referencing a namespace, the value of **Name** will be the property name referenced by namespace. For more information on referencing properties by namespace, see [Referencing Properties by Namespace](../outlook/How-to/Navigation/referencing-properties-by-namespace.md).


## See also


[Column Object](Outlook.Column.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]