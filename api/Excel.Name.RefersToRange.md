---
title: Name.RefersToRange property (Excel)
keywords: vbaxl10.chm490088
f1_keywords:
- vbaxl10.chm490088
ms.prod: excel
api_name:
- Excel.Name.RefersToRange
ms.assetid: 81c0e2fe-8ce6-0df9-9ffa-0931b87487e7
ms.date: 05/01/2019
localization_priority: Normal
---


# Name.RefersToRange property (Excel)

Returns the **[Range](Excel.Range(object).md)** object referred to by a **Name** object. Read-only.


## Syntax

_expression_.**RefersToRange**

_expression_ A variable that represents a **[Name](Excel.Name.md)** object.


## Remarks

If the **Name** object doesn't refer to a range (for example, if it refers to a constant or a formula), this property fails.

To change the range that a name refers to, use the **[RefersTo](Excel.Name.RefersTo.md)** property.


## Example

This example displays the number of rows and columns in the print area on the active worksheet.

> [!NOTE] 
> Ensure that you establish a print area on the active sheet of the current workbook.

```vb
p = Sheets(ActiveSheet.Name).Names("Print_Area").RefersToRange.Value 
MsgBox "Print_Area: " & UBound(p, 1) & " rows, " & _ 
 UBound(p, 2) & " columns"
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
