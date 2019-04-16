---
title: Slicer.Copy method (Excel)
keywords: vbaxl10.chm905091
f1_keywords:
- vbaxl10.chm905091
ms.prod: excel
api_name:
- Excel.Slicer.Copy
ms.assetid: 265e7819-db8b-deab-5ab1-2cc9782cd800
ms.date: 06/08/2017
localization_priority: Normal
---


# Slicer.Copy method (Excel)

Copies the specified slicer to the clipboard.


## Syntax

_expression_.**Copy**

_expression_ A variable that represents a '[Slicer](Excel.Slicer.md)' object.


## Example

The following code example accesses the Customer slicer by using the  **[Range](Excel.Shapes.Range.md)** property of the **[Shapes](Excel.Shapes.md)** collection, and then copies and pastes it into the active worksheet.


```vb
ActiveSheet.Shapes.Range(Array("Customer")).Select 
Selection.Copy 
ActiveSheet.Paste 

```

Alternatively, you can perform the same operation by using the  **[Slicers](Excel.SlicerCache.Slicers.md)** property of the **[SlicerCaches](Excel.SlicerCaches.md)** collection to access the slicer, as shown in the following code example.




```vb
ActiveWorkbook.SlicerCaches("Slicer_Customer") _ 
 .Slicers("Customer").Copy 
ActiveSheet.Paste
```


## See also


[Slicer Object](Excel.Slicer.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]