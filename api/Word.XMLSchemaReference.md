---
title: XMLSchemaReference object (Word)
keywords: vbawd10.chm496
f1_keywords:
- vbawd10.chm496
ms.prod: word
api_name:
- Word.XMLSchemaReference
ms.assetid: 54142ef1-f731-3f82-2dc0-809d8a041b73
ms.date: 06/08/2017
localization_priority: Normal
---


# XMLSchemaReference object (Word)

Represents an individual schema that is attached to a document.


## Remarks

Use the  **XMLSchemaReference** property to return an **XMLSchemaReference** object for a **ChildNodeSuggestion** object. The following example inserts the suggested XML child element if the XML schema referenced is the SimpleSample schema.


```vb
Dim objSuggestion As XMLChildNodeSuggestion 
 
For Each objSuggestion In ActiveDocument _ 
 .ChildNodeSuggestions 
 
 If objSuggestion.XMLSchemaReference = "SimpleSample" Then 
 objSuggestion.Insert 
 End If 
 
Next
```


> [!NOTE] 
> The SimpleSample schema is included in the Smart Document Software Development Kit (SDK). For more information, refer to the Smart Document SDK on the Microsoft Developer Network (MSDN) Web site.


## See also


[Word Object Model Reference](overview/Word/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]