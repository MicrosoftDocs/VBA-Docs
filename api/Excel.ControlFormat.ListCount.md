---
title: ControlFormat.ListCount property (Excel)
keywords: vbaxl10.chm630081
f1_keywords:
- vbaxl10.chm630081
ms.prod: excel
api_name:
- Excel.ControlFormat.ListCount
ms.assetid: 9f7b60aa-8bf9-a7ec-c198-0a6f6316cc3c
ms.date: 04/23/2019
localization_priority: Normal
---


# ControlFormat.ListCount property (Excel)

Returns the number of entries in a list box or combo box. Returns 0 (zero) if there are no entries in the list. Read-only **Long**.


## Syntax

_expression_.**ListCount**

_expression_ A variable that represents a **[ControlFormat](Excel.ControlFormat.md)** object.


## Example

This example adjusts a combo box to display all the entries in its list. If `Shapes(1)` does not represent a combo box, this example fails.

```vb
Set cf = Worksheets(1).Shapes(1).ControlFormat 
cf.DropDownLines = cf.ListCount
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]