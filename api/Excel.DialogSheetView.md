---
title: DialogSheetView object (Excel)
keywords: vbaxl10.chm786072
f1_keywords:
- vbaxl10.chm786072
ms.prod: excel
api_name:
- Excel.DialogSheetView
ms.assetid: d468b3e8-c73e-d94a-0902-193f6983d893
ms.date: 03/29/2019
localization_priority: Normal
---


# DialogSheetView object (Excel)

Represents the current **Dialog** sheet view in a workbook.


## Remarks

To access this object, you must have a dialog sheet that was developed in the active workbook. Without the dialog sheet, the view properties for the object return an empty string value.


## Example

The following example turns on the dialog sheet view for the active workbook.

```vb
Worksheets("Sheet1").DialogSheetView.Visible = True
```

## Properties

- [Application](Excel.DialogSheetView.Application.md)
- [Creator](Excel.DialogSheetView.Creator.md)
- [Parent](Excel.DialogSheetView.Parent.md)
- [Sheet](Excel.DialogSheetView.Sheet.md)

## See also

- [Excel Object Model Reference](overview/Excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]