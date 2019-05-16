---
title: Slicer.Cut method (Excel)
keywords: vbaxl10.chm905090
f1_keywords:
- vbaxl10.chm905090
ms.prod: excel
api_name:
- Excel.Slicer.Cut
ms.assetid: a8778661-612f-0031-78b0-d59bb87fdf62
ms.date: 05/16/2019
localization_priority: Normal
---


# Slicer.Cut method (Excel)

Cuts the specified slicer and copies it to the clipboard.


## Syntax

_expression_.**Cut**

_expression_ A variable that represents a **[Slicer](Excel.Slicer.md)** object.


## Example

The following code example accesses the Customer slicer by using the **[Range](Excel.Shapes.Range.md)** property of the **Shapes** collection, and then cuts and pastes it into the active worksheet.

```vb
ActiveSheet.Shapes.Range(Array("Customer")).Select 
Selection.Cut 
ActiveSheet.Paste 

```

<br/>

Alternatively, you can perform the same operation by using the **[Slicers](Excel.SlicerCache.Slicers.md)** property of the **SlicerCache** object to access the slicer, as shown in the following code example.

```vb
ActiveWorkbook.SlicerCaches("Slicer_Customer") _ 
 .Slicers("Customer").Cut 
ActiveSheet.Paste
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]