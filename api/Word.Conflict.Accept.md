---
title: Conflict.Accept method (Word)
keywords: vbawd10.chm78708837
f1_keywords:
- vbawd10.chm78708837
ms.prod: word
api_name:
- Word.Conflict.Accept
ms.assetid: 3367d8cb-c1b1-3037-06d8-44c275fcfa58
ms.date: 06/08/2017
localization_priority: Normal
---


# Conflict.Accept method (Word)

Accepts the user specified conflict change, and removes the conflict.


## Syntax

_expression_. `Accept`

 _expression_ An expression that returns a [Conflict](./Word.Conflict.md) object.


## Return value

Nothing


## Remarks

In a conflict, a user can choose either to keep or to reject the changes they have made to the content where the conflict exists. The **Accept** method keeps the changes that the user has made.


## Example

The following example accepts all of the conflicts in the active document.


```vb
Dim conf As Conflict 
 
For Each conf In ActiveDocument.CoAuthoring.Conflicts 
    conf.Accept 
Next conf
```

Alternatively, you can use the [AcceptAll](Word.Conflicts.AcceptAll.md) method of the [Conflicts](Word.Conflicts.md) collection object to accept all the conflicts in a document, as shown in the following code example.




```vb
ActiveDocument.CoAuthoring.Conflicts.AcceptAll
```


## See also


[Conflict Object](Word.Conflict.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]