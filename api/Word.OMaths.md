---
title: OMaths object (Word)
keywords: vbawd10.chm1867
f1_keywords:
- vbawd10.chm1867
ms.prod: word
api_name:
- Word.OMaths
ms.assetid: 5e185b0f-b0c9-16f8-3056-c1114dadd3e0
ms.date: 06/08/2017
localization_priority: Normal
---


# OMaths object (Word)

A collection of equations. Use the **OMath** object to access individual members of the collection.


## Remarks

Use the **Add** method to create an equation and add it to a document, selection, or range. The following example creates an equation and uses the **BuildUp** method of the **OMath** collection to convert the equation to professional format.


```vb
Dim objRange As Range 
Dim objEq As OMath 
 
Set objRange = Selection.Range 
objRange.Text = "Celsius = (5/9)(Fahrenheit - 32)" 
Set objRange = Selection.OMaths.Add(objRange) 
Set objEq = objRange.OMaths(1) 
objEq.BuildUp
```


## See also


[Word Object Model Reference](overview/Word/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]