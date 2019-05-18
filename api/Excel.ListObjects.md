---
title: ListObjects object (Excel)
keywords: vbaxl10.chm731072
f1_keywords:
- vbaxl10.chm731072
ms.prod: excel
api_name:
- Excel.ListObjects
ms.assetid: 3a888055-1ed0-d37d-0586-ced999dc1c42
ms.date: 03/30/2019
localization_priority: Normal
---


# ListObjects object (Excel)

A collection of all the **[ListObject](Excel.ListObject.md)** objects on a worksheet. Each **ListObject** object represents a table on the worksheet.


## Remarks

Use the **[ListObjects](Excel.Worksheet.ListObjects.md)** property of the **Worksheet** object to return the **ListObjects** collection.


## Example

The following example creates a new **ListObjects** collection that represents all the tables on a worksheet.

```vb
Set myWorksheetLists = Worksheets(1).ListObjects
```


## Methods

- [Add](Excel.ListObjects.Add.md)

## Properties

- [Application](Excel.ListObjects.Application.md)
- [Count](Excel.ListObjects.Count.md)
- [Creator](Excel.ListObjects.Creator.md)
- [Item](Excel.ListObjects.Item.md)
- [Parent](Excel.ListObjects.Parent.md)


## See also

- [Excel Object Model Reference](overview/Excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
