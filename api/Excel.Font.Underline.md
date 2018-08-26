---
title: Font.Underline Property (Excel)
keywords: vbaxl10.chm559086
f1_keywords:
- vbaxl10.chm559086
ms.prod: excel
api_name:
- Excel.Font.Underline
ms.assetid: 81a2bdd2-bebd-b3ca-e0c3-6dd55280fcc0
ms.date: 06/08/2017
---


# Font.Underline Property (Excel)

<<<<<<< HEAD
Returns or sets the type of underline applied to the font. Can be one of the following  **[XlUnderlineStyle](Excel.XlUnderlineStyle.md)** constants. Read/write **Variant** .
=======
Returns or sets the type of underline applied to the font. Read/write **Variant**.
>>>>>>> master


## Syntax

 _expression_. `Underline`

 _expression_ A variable that represents a [Font](./Excel.Font(Graph property).md) object.


## Remarks

<<<<<<< HEAD




| **XlUnderlineStyle** can be one of these **XlUnderlineStyle** constants.|
| **xlUnderlineStyleNone**|
| **xlUnderlineStyleSingle**|
| **xlUnderlineStyleDouble**|
| **xlUnderlineStyleSingleAccounting**|
| **xlUnderlineStyleDoubleAccounting**|
=======
Can be one of the **[XlUnderlineStyle](Excel.XlUnderlineStyle.md)** constants.
>>>>>>> master

## Example

This example sets the font in the active cell on Sheet1 to single underline.


```vb
Worksheets("Sheet1").Activate 
ActiveCell.Font.Underline = xlUnderlineStyleSingle
```


## See also


[Font Object](Excel.Font(object).md)

