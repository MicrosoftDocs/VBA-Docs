---
title: ControlFormat.RemoveAllItems method (Excel)
keywords: vbaxl10.chm630074
f1_keywords:
- vbaxl10.chm630074
ms.prod: excel
api_name:
- Excel.ControlFormat.RemoveAllItems
ms.assetid: de8e1721-45e1-eca9-d35d-7d72c32dc0bf
ms.date: 04/23/2019
localization_priority: Normal
---


# ControlFormat.RemoveAllItems method (Excel)

Removes all entries from a Microsoft Excel list box or combo box.


## Syntax

_expression_.**RemoveAllItems**

_expression_ A variable that represents a **[ControlFormat](Excel.ControlFormat.md)** object.


## Example

This example removes all items from a list box. If `Shapes(2)` doesn't represent a list box, this example fails.

```vb
Worksheets(1).Shapes(2).ControlFormat.RemoveAllItems
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]