---
title: AutoFilter object (Excel)
keywords: vbaxl10.chm537072
f1_keywords:
- vbaxl10.chm537072
ms.prod: excel
api_name:
- Excel.AutoFilter
ms.assetid: 1a6fcf3b-52be-b599-029b-a3c53d12f85e
ms.date: 03/29/2019
localization_priority: Normal
---


# AutoFilter object (Excel)

Represents autofiltering for the specified worksheet.

> [!NOTE] 
> When using **AutoFilter** with dates, the format should be consistent with English date separators ("/") instead of local settings ("."). A valid date would be "2/2/2007", whereas "2.2.2007" is invalid.

> [!NOTE] 
> Working with objects (for example, the **Interior** object) requires adding a reference to an object. You will find more information about assigning an object reference to a variable or property in the **[Set](../language/reference/User-Interface-Help/set-statement.md)** statement.


## Example

Use the **[AutoFilter](Excel.Worksheet.AutoFilter.md)** property of the **Worksheet** object to return the **AutoFilter** object. Use the **Filters** property to return a collection of individual column filters. Use the **Range** property to return the **Range** object that represents the entire filtered range. 

The following example stores the address and filtering criteria for the current filtering, and then applies new filters.

```vb
Dim w As Worksheet 
Dim filterArray() 
Dim currentFiltRange As String 
 
Sub ChangeFilters() 
 
Set w = Worksheets("Crew") 
With w.AutoFilter 
 currentFiltRange = .Range.Address 
 With .Filters 
 ReDim filterArray(1 To .Count, 1 To 3) 
 For f = 1 To .Count 
 With .Item(f) 
 If .On Then 
 filterArray(f, 1) = .Criteria1 
 If .Operator Then 
 filterArray(f, 2) = .Operator 
 filterArray(f, 3) = .Criteria2 
 End If 
 End If 
 End With 
 Next 
 End With 
End With 
 
w.AutoFilterMode = False 
w.Range("A1").AutoFilter field:=1, Criteria1:="S" 
 
End Sub
```

<br/>

To create an **AutoFilter** object for a worksheet, you must turn autofiltering on for a range on the worksheet either manually or by using the **[AutoFilter](Excel.Range.AutoFilter.md)** method of the **Range** object. The following example uses the values stored in module-level variables in the previous example to restore the original autofiltering to the Crew worksheet.

```vb
Sub RestoreFilters() 
Set w = Worksheets("Crew") 
w.AutoFilterMode = False 
For col = 1 To UBound(filterArray(), 1) 
 If Not IsEmpty(filterArray(col, 1)) Then 
 If filterArray(col, 2) Then 
 w.Range(currentFiltRange).AutoFilter field:=col, _ 
 Criteria1:=filterArray(col, 1), _ 
 Operator:=filterArray(col, 2), _ 
 Criteria2:=filterArray(col, 3) 
 Else 
 w.Range(currentFiltRange).AutoFilter field:=col, _ 
 Criteria1:=filterArray(col, 1) 
 End If 
 End If 
Next 
End Sub 

```




## Methods

- [ApplyFilter](Excel.AutoFilter.ApplyFilter.md)
- [ShowAllData](Excel.AutoFilter.ShowAllData.md)

## Properties

- [Application](Excel.AutoFilter.Application.md)
- [Creator](Excel.AutoFilter.Creator.md)
- [FilterMode](Excel.AutoFilter.FilterMode.md)
- [Filters](Excel.AutoFilter.Filters.md)
- [Parent](Excel.AutoFilter.Parent.md)
- [Range](Excel.AutoFilter.Range.md)
- [Sort](Excel.AutoFilter.Sort.md)

## See also

- [Excel Object Model Reference](overview/Excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
