---
title: Filter object (Excel)
keywords: vbaxl10.chm541072
f1_keywords:
- vbaxl10.chm541072
ms.prod: excel
api_name:
- Excel.Filter
ms.assetid: 950023f9-a984-01fa-aa77-947cbbff0433
ms.date: 06/08/2017
---


# Filter object (Excel)

Represents a filter for a single column.


## Remarks

 The **Filter** object is a member of the **[Filters](Excel.Filters.md)** collection. The **Filters** collection contains all the filters in an autofiltered range.


## Example

Use  **[Filters](Excel.AutoFilter.Filters.md)** ( _index_ ), where _index_ is the filter title or index number, to return a single **Filter** object. The following example sets a variable to the value of the **[On](Excel.Filter.On.md)** property of the filter for the first column in the filtered range on the Crew worksheet.


```vb
Set w = Worksheets("Crew") 
If w.AutoFilterMode Then 
 filterIsOn = w.AutoFilter.Filters(1).On 
End If
```

Note that all the properties of the  **Filter** object are read-only. To set these properties, apply autofiltering manually or using the **[AutoFilter](Excel.Range.AutoFilter.md)** method of the **[Range](Excel.Range(object).md)** object, as shown in the following example.




```vb
Set w = Worksheets("Crew") 
w.Cells.AutoFilter field:=2, Criteria1:="Crucial", _ 
 Operator:=xlOr, Criteria2:="Important"
```


## Properties



|**Name**|
|:-----|
|[Application](Excel.Filter.Application.md)|
|[Count](Excel.Filter.Count.md)|
|[Creator](Excel.Filter.Creator.md)|
|[Criteria1](Excel.Filter.Criteria1.md)|
|[Criteria2](Excel.Filter.Criteria2.md)|
|[On](Excel.Filter.On.md)|
|[Operator](Excel.Filter.Operator.md)|
|[Parent](Excel.Filter.Parent.md)|

## See also


[Excel Object Model Reference](overview/Excel/object-model.md)
