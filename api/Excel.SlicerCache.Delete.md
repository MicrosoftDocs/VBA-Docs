---
title: SlicerCache.Delete method (Excel)
keywords: vbaxl10.chm897090
f1_keywords:
- vbaxl10.chm897090
ms.prod: excel
api_name:
- Excel.SlicerCache.Delete
ms.assetid: 34bc2dce-5286-deb2-995d-c64f146a2cd7
ms.date: 06/08/2017
localization_priority: Normal
---


# SlicerCache.Delete method (Excel)

Deletes the specified slicer cache and the slicers associated with it.


## Syntax

_expression_. `Delete`

_expression_ A variable that represents a '[SlicerCache](Excel.SlicerCache.md)' object.


## Remarks

To delete a particular slicer independently of the slicer cache, use the  **[Delete](Excel.Slicer.Delete.md)** method of the **[Slicer](Excel.Slicer.md)** object instead.


## Example

The following code example deletes the  `Slicer_Country` slicer cache and the `Country` slicer associated with that slicer cache.


```vb
ActiveWorkbook.SlicerCaches("Slicer_Country").Delete
```


## See also


[SlicerCache Object](Excel.SlicerCache.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]