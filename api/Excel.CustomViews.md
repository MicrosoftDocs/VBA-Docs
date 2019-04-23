---
title: CustomViews object (Excel)
keywords: vbaxl10.chm505072
f1_keywords:
- vbaxl10.chm505072
ms.prod: excel
api_name:
- Excel.CustomViews
ms.assetid: f970bdf7-371b-ba41-89a3-bef2c6907f1a
ms.date: 03/29/2019
localization_priority: Normal
---


# CustomViews object (Excel)

A collection of custom workbook views.


## Remarks

Each view is represented by a **[CustomView](Excel.CustomView.md)** object.


## Example

Use the **[CustomViews](excel.workbook.customviews.md)** property of the **Workbook** object to return the **CustomViews** collection. 

Use the **Add** method to create a new custom view and add it to the **CustomViews** collection. The following example creates a new custom view named "Summary."

```vb
ActiveWorkbook.CustomViews.Add "Summary", True, True

```


## Methods

- [Add](Excel.CustomViews.Add.md)
- [Item](Excel.CustomViews.Item.md)

## Properties

- [Application](Excel.CustomViews.Application.md)
- [Count](Excel.CustomViews.Count.md)
- [Creator](Excel.CustomViews.Creator.md)
- [Parent](Excel.CustomViews.Parent.md)

## See also

- [Excel Object Model Reference](overview/Excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]