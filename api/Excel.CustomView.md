---
title: CustomView object (Excel)
keywords: vbaxl10.chm507072
f1_keywords:
- vbaxl10.chm507072
ms.prod: excel
api_name:
- Excel.CustomView
ms.assetid: e16b1920-faeb-62d4-4d27-59745c4f5355
ms.date: 03/29/2019
localization_priority: Normal
---


# CustomView object (Excel)

Represents a custom workbook view.


## Remarks

The **CustomView** object is a member of the **[CustomViews](Excel.CustomViews.md)** collection.


## Example

Use **CustomViews** (_index_), where _index_ is the name or index number of the custom view, to return a **CustomView** object. The following example shows the custom view named **Current Inventory**.

```vb
ThisWorkbook.CustomViews("Current Inventory").Show
```


## Methods

- [Delete](Excel.CustomView.Delete.md)
- [Show](Excel.CustomView.Show.md)

## Properties

- [Application](Excel.CustomView.Application.md)
- [Creator](Excel.CustomView.Creator.md)
- [Name](Excel.CustomView.Name.md)
- [Parent](Excel.CustomView.Parent.md)
- [PrintSettings](Excel.CustomView.PrintSettings.md)
- [RowColSettings](Excel.CustomView.RowColSettings.md)


## See also

- [Excel Object Model Reference](overview/Excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]