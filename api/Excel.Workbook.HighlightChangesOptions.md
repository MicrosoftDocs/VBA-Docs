---
title: Workbook.HighlightChangesOptions method (Excel)
keywords: vbaxl10.chm199172
f1_keywords:
- vbaxl10.chm199172
ms.prod: excel
api_name:
- Excel.Workbook.HighlightChangesOptions
ms.assetid: ac69ee3e-c5ea-5ac0-418a-0b94d56a8777
ms.date: 06/08/2017
localization_priority: Normal
---


# Workbook.HighlightChangesOptions method (Excel)

Controls how changes are shown in a shared workbook.


## Syntax

_expression_. `HighlightChangesOptions`( `_When_` , `_Who_` , `_Where_` )

_expression_ A variable that represents a [Workbook](./Excel.Workbook.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _When_|Optional| **Variant**|The changes that are shown. Can be one of the following  **[xlHighlightChangesTime](Excel.XlHighlightChangesTime.md)** constants: **xlSinceMyLastSave** , **xlAllChanges** , or **xlNotYetReviewed**.|
| _Who_|Optional| **Variant**|The user or users whose changes are shown. Can be "Everyone," "Everyone but Me," or the name of one of the users of the shared workbook.|
| _Where_|Optional| **Variant**|An A1-style range reference that specifies the area to check for changes.|

## Example

This example shows changes to the shared workbook on a separate worksheet.


```vb
With ActiveWorkbook 
 .HighlightChangesOptions _ 
 When:=xlSinceMyLastSave, _ 
 Who:="Everyone" 
 .ListChangesOnNewSheet = True 
End With 

```


## See also


[Workbook Object](Excel.Workbook.md)

