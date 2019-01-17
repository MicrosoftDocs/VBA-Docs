---
title: Areas object (Excel)
keywords: vbaxl10.chm196072
f1_keywords:
- vbaxl10.chm196072
ms.prod: excel
api_name:
- Excel.Areas
ms.assetid: 43d05ef3-7ae2-2881-dec2-6fec8281f045
ms.date: 06/08/2017
localization_priority: Priority
---


# Areas object (Excel)

A collection of the areas, or contiguous blocks of cells, within a selection. 


## Remarks

There's no singular Area object; individual members of the  **Areas** collection are **[Range](Excel.Range(object).md)** objects. The **Areas** collection contains one **Range** object for each discrete, contiguous range of cells within the selection. If the selection contains only one area, the **Areas** collection contains a single **Range** object that corresponds to that selection.


## Example

Use the  **Areas** property to return the **Areas** collection. The following example clears the current selection if it contains more than one area.


```vb
If Selection.Areas.Count <> 1 Then Selection.Clear
```

Use  **Areas** ( _index_ ), where _index_ is the area index number, to return a single **Range** object from the collection. The index numbers correspond to the order in which the areas were selected. The following example clears the first area in the current selection if the selection contains more than one area.




```vb
If Selection.Areas.Count <> 1 Then 
 Selection.Areas(1).Clear 
End If
```

Some operations cannot be performed on more than one area in a selection at the same time; you must loop through the individual areas in the selection and perform the operations on each area separately. The following example performs the operation named "myOperation" on the selected range if the selection contains only one area; if the selection contains multiple areas, the example performs myOperation on each individual area in the selection.




```vb
Set rangeToUse = Selection 
If rangeToUse.Areas.Count = 1 Then 
 myOperation rangeToUse 
Else 
 For Each singleArea in rangeToUse.Areas 
 myOperation singleArea 
 Next 
End If
```


## Properties



|Name|
|:-----|
|[Application](Excel.Areas.Application.md)|
|[Count](Excel.Areas.Count.md)|
|[Creator](Excel.Areas.Creator.md)|
|[Item](Excel.Areas.Item.md)|
|[Parent](Excel.Areas.Parent.md)|

## See also


[Excel Object Model Reference](overview/Excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]