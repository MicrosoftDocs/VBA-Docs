---
title: Columns.Remove method (Outlook)
keywords: vbaol11.chm2742
f1_keywords:
- vbaol11.chm2742
ms.prod: outlook
api_name:
- Outlook.Columns.Remove
ms.assetid: f567879c-f37a-2b65-b4a5-832b6f3acdf8
ms.date: 06/08/2017
localization_priority: Normal
---


# Columns.Remove method (Outlook)

Removes the  **[Column](Outlook.Column.md)** object specified by _Index_ and resets the **[Table](Outlook.Table.md)**.


## Syntax

_expression_.**Remove** (_Index_)

_expression_ A variable that represents a '[Columns](Outlook.Columns.md)' object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Index_|Required| **Variant**|A 1-based index value that can be either a **Long** representing the column index for the **Columns** collection or a **String** representing the **[Name](Outlook.Column.Name.md)** of the **Column**.|

## Remarks

The **Remove** method resets the **Table** by moving the current row to just before the first row of the **Table**. If, however, an invalid _Index_ has been specified, then it will not remove any column or reset the **Table**.

Returns an error message if an invalid  _Index_ has been specified.


## See also


[Columns Object](Outlook.Columns.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]