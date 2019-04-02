---
title: Workbooks object (Excel)
keywords: vbaxl10.chm202072
f1_keywords:
- vbaxl10.chm202072
ms.prod: excel
api_name:
- Excel.Workbooks
ms.assetid: f768da57-013a-e652-0f5d-60b03aa4240a
ms.date: 04/03/2019
localization_priority: Normal
---


# Workbooks object (Excel)

A collection of all the **[Workbook](Excel.Workbook.md)** objects that are currently open in the Microsoft Excel application.


## Remarks

For more information about using a single **Workbook** object, see the **[Workbook](Excel.Workbook.md)** object.


## Example

Use the **[Workbooks](Excel.Application.Workbooks.md)** property of the **Application** object to return the **Workbooks** collection. The following example closes all open workbooks.

```vb
Workbooks.Close
```

<br/>

Use the **Add** method to create a new, empty workbook and add it to the collection. The following example adds a new, empty workbook to Microsoft Excel.

```vb
Workbooks.Add
```

<br/>

Use the **Open** method to open a file. This creates a new workbook for the opened file. The following example opens the file Array.xls as a read-only workbook.

```vb
Workbooks.Open FileName:="Array.xls", ReadOnly:=True
```


## Methods

- [Add](Excel.Workbooks.Add.md)
- [CanCheckOut](Excel.Workbooks.CanCheckOut.md)
- [CheckOut](Excel.Workbooks.CheckOut.md)
- [Close](Excel.Workbooks.Close.md)
- [Open](Excel.Workbooks.Open.md)
- [OpenDatabase](Excel.Workbooks.OpenDatabase.md)
- [OpenText](Excel.Workbooks.OpenText.md)
- [OpenXML](Excel.Workbooks.OpenXML.md)

## Properties

- [Application](Excel.Workbooks.Application.md)
- [Count](Excel.Workbooks.Count.md)
- [Creator](Excel.Workbooks.Creator.md)
- [Item](Excel.Workbooks.Item.md)
- [Parent](Excel.Workbooks.Parent.md)


## See also

- [Excel Object Model Reference](overview/Excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
