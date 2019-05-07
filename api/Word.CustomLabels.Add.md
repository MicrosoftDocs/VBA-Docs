---
title: CustomLabels.Add method (Word)
keywords: vbawd10.chm152436837
f1_keywords:
- vbawd10.chm152436837
ms.prod: word
api_name:
- Word.CustomLabels.Add
ms.assetid: 12bfd8d5-ab6e-7946-563c-0bb9c21393c9
ms.date: 06/08/2017
localization_priority: Normal
---


# CustomLabels.Add method (Word)

Adds a custom mailing label to the  **CustomLabels** collection. Returns a **CustomLabel** object that represents the custom mailing label.


## Syntax

_expression_.**Add** (_Name_, _DotMatrix_)

_expression_ Required. A variable that represents a '[CustomLabels](Word.customlabels.md)' collection.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Name_|Required| **String**|The name for the custom mailing labels.|
| _DotMatrix_|Optional| **Variant**| **True** to have the mailing labels printed on a dot-matrix printer.|

## Return value

CustomLabel


## Example

This example adds a custom mailing label named Return Address, sets the page size, and then creates a page of these labels.


```vb
Sub ReturnAddrLabel() 
 Dim ml As CustomLabel 
 Dim addr As String 
 
 Set ml = Application.MailingLabel.CustomLabels _ 
 .Add(Name:="Return Address", DotMatrix:=False) 
 ml.PageSize = wdCustomLabelLetter 
 addr = "Dave Edson" & vbCr & "123 Skye St." & vbCr _ 
 & "Our Town, WA 98004" 
 Application.MailingLabel.CreateNewDocument _ 
 Name:="Return Address", Address:=addr, ExtractAddress:=False 
End Sub
```


## See also


[CustomLabels Collection Object](Word.customlabels.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]