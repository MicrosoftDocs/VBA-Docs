---
title: ListFormat.ApplyListTemplateWithLevel method (Word)
keywords: vbawd10.chm163578072
f1_keywords:
- vbawd10.chm163578072
ms.prod: word
api_name:
- Word.ListFormat.ApplyListTemplateWithLevel
ms.assetid: 53d107d1-7a6c-b94c-19b9-2794e20ef1cb
ms.date: 06/08/2017
localization_priority: Normal
---


# ListFormat.ApplyListTemplateWithLevel method (Word)

Applies a set of list-formatting characteristics, optionally for a specified level.


## Syntax

_expression_. `ApplyListTemplateWithLevel`( `_ListTemplate_` , `_ContinuePreviousList_` , `_ApplyTo_` , `_DefaultListBehavior_` , `_ApplyLevel_` )

_expression_ A variable that represents a '[ListFormat](Word.ListFormat.md)' object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _ListTemplate_|Required| **[ListTemplate](Word.ListTemplate.md)**|The list template to be applied.|
| _ContinuePreviousList_|Optional| **Variant**| **True** to continue the numbering from the previous list; **False** to start a new list.|
| _ApplyTo_|Optional| **Variant**|One of the **[WdListApplyTo](Word.WdListApplyTo.md)** constants that specifies the portion of the list that the list template will be applied to.|
| _DefaultListBehavior_|Optional| **Variant**|Sets a value that specifies whether Microsoft Word uses new Web-oriented formatting for better list display. Can be either of the following  **[WdDefaultListBehavior](Word.WdDefaultListBehavior.md)** constants: **wdWord8ListBehavior** (use formatting compatible with Microsoft Word 97) or **wdWord9ListBehavior** (use Web-oriented formatting). For compatibility reasons, the default constant is **wdWord8ListBehavior**, but in new procedures you should use **wdWord9ListBehavior** to take advantage of improved Web-oriented formatting for indenting and multiple-level lists.|
| _ApplyLevel_|Optional| **Variant**|The level to which the list template is to be applied.|

## Example

The following example sets the variable  `myRange` to a range in the active document, and then it verifies whether the range has list formatting. If no list formatting has been applied, the fourth outline-numbered list template is applied to the range.


```vb
Set myDoc = ActiveDocument 
Set myRange = myDoc.Range( _ 
 Start:= myDoc.Paragraphs(3).Range.Start, _ 
 End:=myDoc.Paragraphs(6).Range.End) 
If myRange.ListFormat.ListType = wdListNoNumbering Then 
 myRange.ListFormat.ApplyListTemplate _ 
 ListTemplate:=ListGalleries(wdOutlineNumberGallery) _ 
 .ListTemplates(4) 
End If
```


## See also


[ListFormat Object](Word.ListFormat.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]