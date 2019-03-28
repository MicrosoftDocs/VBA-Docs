---
title: Filters object (Excel)
keywords: vbaxl10.chm539072
f1_keywords:
- vbaxl10.chm539072
ms.prod: excel
api_name:
- Excel.Filters
ms.assetid: a714ed69-7772-5ade-3acd-f3e3d98db62c
ms.date: 03/29/2019
localization_priority: Normal
---


# Filters object (Excel)

A collection of **[Filter](Excel.Filter.md)** objects that represents all the filters in an autofiltered range.


## Example

Use the **[Filters](Excel.AutoFilter.Filters.md)** property of the **AutoFilter** object to return the **Filters** collection. The following example creates a list that contains the criteria and operators for the filters in the autofiltered range on the Crew worksheet.

```vb
Dim f As Filter 
Dim w As Worksheet 
Const ns As String = "Not set" 
 
Set w = Worksheets("Crew") 
Set w2 = Worksheets("FilterData") 
rw = 1 
For Each f In w.AutoFilter.Filters 
 If f.On Then 
 c1 = Right(f.Criteria1, Len(f.Criteria1) - 1) 
 If f.Operator Then 
 op = f.Operator 
 c2 = Right(f.Criteria2, Len(f.Criteria2) - 1) 
 Else 
 op = ns 
 c2 = ns 
 End If 
 Else 
 c1 = ns 
 op = ns 
 c2 = ns 
 End If 
 w2.Cells(rw, 1) = c1 
 w2.Cells(rw, 2) = op 
 w2.Cells(rw, 3) = c2 
 rw = rw + 1 
Next
```

<br/>

Use **Filters** (_index_), where _index_ is the filter title or index number, to return a single **Filter** object. The following example sets a variable to the value of the **On** property of the filter for the first column in the filtered range on the Crew worksheet.

```vb
Set w = Worksheets("Crew") 
If w.AutoFilterMode Then 
 filterIsOn = w.AutoFilter.Filters(1).On 
End If
```


## Properties

- [Application](Excel.Filters.Application.md)
- [Count](Excel.Filters.Count.md)
- [Creator](Excel.Filters.Creator.md)
- [Item](Excel.Filters.Item.md)
- [Parent](Excel.Filters.Parent.md)


## See also

- [Excel Object Model Reference](overview/Excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]