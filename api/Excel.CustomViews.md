---
title: CustomViews Object (Excel)
keywords: vbaxl10.chm505072
f1_keywords:
- vbaxl10.chm505072
ms.prod: excel
api_name:
- Excel.CustomViews
ms.assetid: f970bdf7-371b-ba41-89a3-bef2c6907f1a
ms.date: 06/08/2017
---


# CustomViews Object (Excel)

A collection of custom workbook views.


## Remarks

 Each view is represented by a **[CustomView](Excel.CustomView.md)** object.


## Example

Use the  **CustomViews** property to return the **CustomViews** collection. Use the **[Add](Excel.CustomViews.Add.md)** method to create a new custom view and add it to the **CustomViews** collection. The following example creates a new custom view named "Summary."


```vb
ActiveWorkbook.CustomViews.Add "Summary", True, True
```


## See also


[Excel Object Model Reference](overview/Excel/object-model.md)


