---
title: MailMergeFields.AddAsk method (Word)
keywords: vbawd10.chm153026662
f1_keywords:
- vbawd10.chm153026662
ms.prod: word
api_name:
- Word.MailMergeFields.AddAsk
ms.assetid: ea52714b-c7c3-a175-67b3-3ce9645218d2
ms.date: 06/08/2017
localization_priority: Normal
---


# MailMergeFields.AddAsk method (Word)

Adds an ASK field to a mail merge main document. Returns a  **MailMergeField** object.


## Syntax

_expression_. `AddAsk`( `_Range_` , `_Name_` , `_Prompt_` , `_DefaultAskText_` , `_AskOnce_` )

_expression_ Required. A variable that represents a '[MailMergeFields](Word.mailmergefields.md)' collection.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Range_|Required| **Range object**|The location for the ASK field.|
| _Name_|Required| **String**|The bookmark name that the response or default text is assigned to. Use a REF field with the bookmark name to display the result in a document.|
| _Prompt_|Optional| **Variant**|The text that's displayed in the dialog box.|
| _DefaultAskText_|Optional| **Variant**|The default response, which appears in the text box when the dialog box is displayed. Corresponds to the \d switch for an ASK field.|
| _AskOnce_|Optional| **Variant**| **True** to display the dialog box only once instead of each time a new record is merged. Corresponds to the \o switch for an ASK field.|

## Return value

MailMergeField


## Remarks

When updated, an ASK field displays a dialog box that prompts you for text to assign to the specified bookmark.


## Example

This example adds an ASK field at the end of the active mail merge main document.


```vb
Dim rngTemp As Range 
 
Set rngTemp = ActiveDocument.Content 
 
rngTemp.Collapse Direction:=wdCollapseEnd 
ActiveDocument.MailMerge.Fields.AddAsk _ 
 Range:=rngTemp, _ 
 Prompt:="Type your company name", _ 
 Name:="company", AskOnce:=True
```

This example adds an ASK field after the last mail merge field in Main.doc.




```vb
Dim colMailMergeFields As Object 
Dim rngTemp As Range 
 
Set colMailMergeFields = Documents("Main.doc").MailMerge.Fields 
 
colMailMergeFields(colMailMergeFields.Count).Select 
 
Set rngTemp = Selection.Range 
 
rngTemp.Collapse wdCollapseEnd 
colMailMergeFields.AddAsk Range:=rngTemp, Name:="name", _ 
 Prompt:="What is your name"
```


## See also


[MailMergeFields Collection Object](Word.mailmergefields.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]