---
title: TableStyle.Condition method (Word)
keywords: vbawd10.chm244776976
f1_keywords:
- vbawd10.chm244776976
ms.prod: word
api_name:
- Word.TableStyle.Condition
ms.assetid: f0adb8b7-434d-3134-38d0-d21d221a27d3
ms.date: 06/08/2017
localization_priority: Normal
---


# TableStyle.Condition method (Word)

Returns a  **[ConditionalStyle](Word.ConditionalStyle.md)** object that represents special style formatting for a portion of a table.


## Syntax

_expression_. `Condition`( `_ConditionCode_` )

_expression_ Required. A variable that represents a '[TableStyle](Word.TableStyle.md)' object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _ConditionCode_|Required| [**WdConditionCode**](Word.WdConditionCode.md)|The area of the table to which to apply the formatting.|

## Example

This example selects the first table in the active document and adds a 20 percent shading to odd-numbered columns.


```vb
Sub TableStylesTest() 
 With ActiveDocument 
 
 'Select the table to which the conditional 
 'formatting will apply 
 .Tables(1).Select 
 
 'Specify the conditional formatting 
 .Styles("Table Grid").Table _ 
 .Condition(wdOddColumnBanding).Shading _ 
 .BackgroundPatternColor = wdColorGray20 
 End With 
End Sub
```


## See also


[TableStyle Object](Word.TableStyle.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]