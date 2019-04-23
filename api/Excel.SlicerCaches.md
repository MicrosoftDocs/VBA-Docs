---
title: SlicerCaches object (Excel)
keywords: vbaxl10.chm894072
f1_keywords:
- vbaxl10.chm894072
ms.prod: excel
api_name:
- Excel.SlicerCaches
ms.assetid: d6097f70-cdc7-3be7-575c-cf43a0765e10
ms.date: 04/02/2019
localization_priority: Normal
---


# SlicerCaches object (Excel)

Represents the collection of slicer caches associated with the specified workbook.


## Remarks

Use the **Item** property of the **SlicerCaches** collection to return a **[SlicerCache](Excel.SlicerCache.md)** object associated with the specified **[Workbook](Excel.Workbook.md)** object. A **SlicerCache** object can be retrieved by using either the value of the **[Index](Excel.SlicerCache.Index.md)** property or the **[Name](Excel.SlicerCache.Name.md)** property of the specified object.


## Example

The following code example retrieves the **SlicerCache** object that represents the slicer cache associated with the Country slicer.

```vb
ActiveWorkbook.SlicerCaches("Slicer_Country")
```


## Methods

- [Add](Excel.SlicerCaches.Add.md)

## Properties

- [Application](Excel.SlicerCaches.Application.md)
- [Count](Excel.SlicerCaches.Count.md)
- [Creator](Excel.SlicerCaches.Creator.md)
- [Item](Excel.SlicerCaches.Item.md)
- [Parent](Excel.SlicerCaches.Parent.md)


## See also

- [Excel Object Model Reference](overview/Excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]