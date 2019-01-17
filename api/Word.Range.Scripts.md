---
title: Range.Scripts property (Word)
keywords: vbawd10.chm157155653
f1_keywords:
- vbawd10.chm157155653
ms.prod: word
api_name:
- Word.Range.Scripts
ms.assetid: 233acf3a-3151-f4f2-e5df-815edeca1dd1
ms.date: 06/08/2017
localization_priority: Normal
---


# Range.Scripts property (Word)

Returns a  **Scripts** collection that represents the collection of HTML scripts in the specified object.


## Syntax

 _expression_. `Scripts`

 _expression_ Required. A variable that represents a '[Range](Word.Range.md)' object.


## Example

This example tests the second  **Script** object in the specified range to determine its language.


```vb
Select Case Selection.Range.Scripts(2).Language 
 Case msoScriptLanguageASP 
 MsgBox "Active Server Pages" 
 Case msoScriptLanguageVisualBasic 
 MsgBox "VBScript" 
 Case msoScriptLanguageJava 
 MsgBox "JavaScript" 
 Case msoScriptLanguageOther 
 MsgBox "Unknown type of script" 
End Select
```


## See also


[Range Object](Word.Range.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]