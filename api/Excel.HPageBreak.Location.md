---
title: HPageBreak.Location property (Excel)
keywords: vbaxl10.chm159078
f1_keywords:
- vbaxl10.chm159078
ms.prod: excel
api_name:
- Excel.HPageBreak.Location
ms.assetid: 7f0ce2ba-21e6-4dc9-8957-ade679aeeabb
ms.date: 04/26/2019
localization_priority: Normal
---


# HPageBreak.Location property (Excel)

Returns or sets the cell (a **[Range](Excel.Range(object).md)** object) that defines the page-break location. Horizontal page breaks are aligned with the top edge of the location cell. Read/write **Range**.


## Syntax

_expression_.**Location** 

_expression_ A variable that represents an **[HPageBreak](Excel.HPageBreak.md)** object.


## Example

This example sets the horizontal page-break location. Note that you must be in Page Break Preview mode in order to set it.

```vb
Set Worksheets(1).HPageBreaks(1).Location = Worksheets(1).Range("e5")
```


> [!NOTE] 
> The **Location** property can only be used to set the horizontal page-break location. To change the location of a **VPageBreak**, you must use **[VPageBreak.Dragoff](Excel.VPageBreak.DragOff.md)**.



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]