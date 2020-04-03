---
title: Columns.Item method (Outlook)
keywords: vbaol11.chm2740
f1_keywords:
- vbaol11.chm2740
ms.prod: outlook
api_name:
- Outlook.Columns.Item
ms.assetid: d9abb503-32ea-d98b-bc43-d818c8b72883
ms.date: 06/08/2017
localization_priority: Normal
---


# Columns.Item method (Outlook)

Obtains a  **[Column](Outlook.Column.md)** object specified by _Index_.


## Syntax

_expression_.**Item** (_Index_)

_expression_ A variable that represents a '[Columns](Outlook.Columns.md)' object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Index_|Required| **Variant**|A 1-based index value that can be either a  **Long** representing the column index for the **Columns** collection or a **String** representing the **[Name](Outlook.Column.Name.md)** of the **Column**.|

## Return value

 A **Column** object that represents the column matching the _Index_ in the **[Table](Outlook.Table.md)**. Returns the error, "Array index out of bounds" if _Index_ is an invalid **Long** integer. Returns **Null** (**Nothing** in Visual Basic) if _Index_ is a **String** representing a column name that cannot be found in the **Table**.


## See also


[Columns Object](Outlook.Columns.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]