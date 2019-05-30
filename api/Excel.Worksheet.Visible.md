---
title: Worksheet.Visible property (Excel)
keywords: vbaxl10.chm174097
f1_keywords:
- vbaxl10.chm174097
ms.prod: excel
api_name:
- Excel.Worksheet.Visible
ms.assetid: 48860564-6079-932e-2cae-0802235be61e
ms.date: 05/30/2019
localization_priority: Normal
---


# Worksheet.Visible property (Excel)

Returns or sets an **[XlSheetVisibility](Excel.XlSheetVisibility.md)** value that determines whether the object is visible.


## Syntax

_expression_.**Visible**

_expression_ A variable that represents a **[Worksheet](Excel.Worksheet.md)** object.


## Example

This example hides Sheet1.

```vb
Worksheets("Sheet1").Visible = False
```

<br/>

This example makes Sheet1 visible.

```vb
Worksheets("Sheet1").Visible = True
```

<br/>

This example makes every sheet in the active workbook visible.

```vb
For Each sh In Sheets 
 sh.Visible = True 
Next sh
```

<br/>

This example creates a new worksheet and then sets its **Visible** property to **xlSheetVeryHidden**. To refer to the sheet, use its object variable, `newSheet`, as shown in the last line of the example. To use the `newSheet` object variable in another procedure, you must declare it as a public variable (`Public newSheet As Object`) in the first line of the module preceding any **Sub** or **Function** procedure.

```vb
Set newSheet = Worksheets.Add 
newSheet.Visible = xlSheetVeryHidden 
newSheet.Range("A1:D4").Formula = "=RAND()"
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
