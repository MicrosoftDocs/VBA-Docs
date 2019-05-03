---
title: Sort object (Excel)
keywords: vbaxl10.chm846072
f1_keywords:
- vbaxl10.chm846072
ms.prod: excel
api_name:
- Excel.Sort
ms.assetid: 637ee681-743c-5196-2bfc-4a5bea025295
ms.date: 04/02/2019
localization_priority: Normal
---


# Sort object (Excel)

Represents a sort of a range of data.


## Example

The following procedure builds and sorts data in a range on the active worksheet.

```vb
Sub SortData() 
 
 'Building data to sort on the active sheet. 
 Range("A1").Value = "Name" 
 Range("A2").Value = "Bill" 
 Range("A3").Value = "Rod" 
 Range("A4").Value = "John" 
 Range("A5").Value = "Paddy" 
 Range("A6").Value = "Kelly" 
 Range("A7").Value = "William" 
 Range("A8").Value = "Janet" 
 Range("A9").Value = "Florence" 
 Range("A10").Value = "Albert" 
 Range("A11").Value = "Mary" 
 MsgBox "The list is out of order. Hit Ok to continue...", vbInformation 
 
 'Selecting a cell within the range. 
 Range("A2").Select 
 
 'Applying sort. 
 With ActiveWorkbook.Worksheets(ActiveSheet.Name).Sort 
 .SortFields.Clear 
 .SortFields.Add Key:=Range("A2:A11"), _ 
 SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal 
 .SetRange Range("A1:A11") 
 .Header = xlYes 
 .MatchCase = False 
 .Orientation = xlTopToBottom 
 .SortMethod = xlPinYin 
 .Apply 
 End With 
 MsgBox "Sort complete.", vbInformation 
 
End Sub
```


## Methods

- [Apply](Excel.Sort.Apply.md)
- [SetRange](Excel.Sort.SetRange.md)

## Properties

- [Application](Excel.Sort.Application.md)
- [Creator](Excel.Sort.Creator.md)
- [Header](Excel.Sort.Header.md)
- [MatchCase](Excel.Sort.MatchCase.md)
- [Orientation](Excel.Sort.Orientation.md)
- [Parent](Excel.Sort.Parent.md)
- [Rng](Excel.Sort.Rng.md)
- [SortFields](Excel.Sort.SortFields.md)
- [SortMethod](Excel.Sort.SortMethod.md)


## See also

- [Excel Object Model Reference](overview/Excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
