---
title: Worksheet.Copy method (Excel)
keywords: vbaxl10.chm174074
f1_keywords:
- vbaxl10.chm174074
ms.prod: excel
api_name:
- Excel.Worksheet.Copy
ms.assetid: ace07575-34f4-a4ae-0922-a2671f2df1ba
ms.date: 05/30/2019
localization_priority: Normal
---


# Worksheet.Copy method (Excel)

Copies the sheet to another location in the current workbook or a new workbook.


## Syntax

_expression_.**Copy** (_Before_, _After_)

_expression_ A variable that represents a **[Worksheet](Excel.Worksheet.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Before_|Optional| **Variant**|The sheet before which the copied sheet will be placed. You cannot specify _Before_ if you specify _After_.|
| _After_|Optional| **Variant**|The sheet after which the copied sheet will be placed. You cannot specify _After_ if you specify _Before_.|

## Remarks

If you don't specify either _Before_ or _After_, Microsoft Excel creates a new workbook that contains the copied **Worksheet** object. The newly created workbook holds the **[Application.ActiveWorkbook](Excel.Application.ActiveWorkbook.md)** property and contains a single worksheet. The single worksheet retains the **[Name](Excel.Worksheet.Name.md)** and **[CodeName](Excel.Worksheet.CodeName.md)** properties of the source worksheet. If the copied worksheet held a worksheet code sheet in a VBA project, that is also carried into the new workbook.

An array selection of multiple worksheets can be copied to a new blank **[Workbook](Excel.Workbook.md)** object in a similar manner.

Source and Destination must be in the same Excel.Application instance, otherwise it will raise a runtime error '1004': No such interface supported, if something like `Sheet1.Copy objWb.Sheets(1)` was used, or a runtime error '1004': Copy method of Worksheet class failed, if something like `ThisWorkbook.Worksheets("Sheet1").Copy objWb.Sheets(1)` was used.

## Example

This example copies Sheet1, placing the copy after Sheet3.

```vb
Worksheets("Sheet1").Copy After:=Worksheets("Sheet3")
```

<br/>

This example first copies Sheet1 to a new blank workbook, and then saves and closes the new workbook.

```vb
Worksheets("Sheet1").Copy
With ActiveWorkbook 
     .SaveAs Filename:=Environ("TEMP") & "\New1.xlsx", FileFormat:=xlOpenXMLWorkbook
     .Close SaveChanges:=False
End With

```

<br/>

This example copies worksheets Sheet1, Sheet2, and Sheet4 to a new blank workbook, and then saves and closes the new workbook.

```vb
Worksheets(Array("Sheet1", "Sheet2", "Sheet4")).Copy
With ActiveWorkbook
     .SaveAs Filename:=Environ("TEMP") & "\New3.xlsx", FileFormat:=xlOpenXMLWorkbook 
     .Close SaveChanges:=False 
End With 

```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
