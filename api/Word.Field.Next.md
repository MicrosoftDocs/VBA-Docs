---
title: Field.Next property (Word)
keywords: vbawd10.chm154075142
f1_keywords:
- vbawd10.chm154075142
ms.prod: word
api_name:
- Word.Field.Next
ms.assetid: be828737-6ac4-9986-4b57-187a7198898d
ms.date: 06/08/2017
localization_priority: Normal
---


# Field.Next property (Word)

Returns the next object in the collection. Read-only.


## Syntax

_expression_.**Next**

_expression_ A variable that represents a **[Field](Word.Field.md)** object.


## Example

This example updates the fields in the first section in the active document as long as the **Next** method returns a **Field** object and the field isn't a FILLIN field.

```vb
If ActiveDocument.Sections(1).Range.Fields.Count >= 1 Then 
 Set myField = ActiveDocument.Fields(1) 
 While Not (myField Is Nothing) 
 If myField.Type <> wdFieldFillIn Then myField.Update 
 Set myField = myField.Next 
 Wend 
End If
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]