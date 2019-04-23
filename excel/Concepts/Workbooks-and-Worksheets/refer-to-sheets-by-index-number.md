---
title: Refer to Sheets by Index Number
keywords: vbaxl10.chm5204441
f1_keywords:
- vbaxl10.chm5204441
ms.prod: excel
ms.assetid: dc947b43-8e96-733a-72e8-3487a4ad9e96
ms.date: 06/08/2017
localization_priority: Normal
---


# Refer to Sheets by Index Number

An index number is a sequential number assigned to a sheet, based on the position of its sheet tab (counting from the left) among sheets of the same type. The following procedure uses the  **[Worksheets](../../../api/Excel.Workbook.Worksheets.md)** property to activate the first worksheet in the active workbook.


```vb
Sub FirstOne() 
 Worksheets(1).Activate 
End Sub
```


If you want to work with all types of sheets (worksheets, charts, modules, and dialog sheets), use the  **[Sheets](../../../api/Excel.Workbook.Sheets.md)** property. The following procedure activates sheet four in the workbook.




```vb
Sub FourthOne() 
 Sheets(4).Activate 
End Sub
```


 **Note**  The index order can change if you move, add, or delete sheets.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
