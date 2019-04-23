---
title: Using With statements (VBA)
keywords: vbcn6.chm1076687
f1_keywords:
- vbcn6.chm1076687
ms.prod: office
ms.assetid: ae7f6296-f151-1a1d-a273-a4b80b18b367
ms.date: 12/26/2018
localization_priority: Normal
---


# Using With statements

The **[With](../../reference/user-interface-help/with-statement.md)** statement lets you specify an [object](../../Glossary/vbe-glossary.md#object) or [user-defined type](../../Glossary/vbe-glossary.md#user-defined-type) once for an entire series of [statements](../../Glossary/vbe-glossary.md#statement). **With** statements make your procedures run faster and help you avoid repetitive typing.

The following example fills a range of cells with the number 30, applies bold formatting, and sets the interior color of the cells to yellow.

```vb
Sub FormatRange() 
 With Worksheets("Sheet1").Range("A1:C10") 
 .Value = 30 
 .Font.Bold = True 
 .Interior.Color = RGB(255, 255, 0) 
 End With 
End Sub
```

You can nest **With** statements for greater efficiency. The following example inserts a formula into cell A1, and then formats the font.

```vb
Sub MyInput() 
 With Workbooks("Book1").Worksheets("Sheet1").Cells(1, 1) 
 .Formula = "=SQRT(50)" 
 With .Font 
 .Name = "Arial" 
 .Bold = True 
 .Size = 8 
 End With 
 End With 
End Sub
```

## See also

- [Visual Basic conceptual topics](../../reference/user-interface-help/visual-basic-conceptual-topics.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]