---
title: OMaths.Add method (Word)
keywords: vbawd10.chm122355914
f1_keywords:
- vbawd10.chm122355914
ms.prod: word
api_name:
- Word.OMaths.Add
ms.assetid: d1372381-e9b3-b551-41ff-fa908800d683
ms.date: 06/08/2017
localization_priority: Normal
---


# OMaths.Add method (Word)

Creates an equation, from the text equation contained within the specified range, and returns a  **Range** object that contains the new equation.


## Syntax

_expression_.**Add** (_Range_)

 _expression_ An expression that returns an [OMaths](./Word.OMaths.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Range_|Required| **Range**|Specifies a range that contains a text equation.|

## Return value

Range


## Example

The following example inserts an equation into the document at the cursor or replacing the selected text.


```vb
Dim objRange As Range 
Dim objEq As OMath 
 
Set objRange = Selection.Range 
objRange.Text = "Celsius = (5/9)(Fahrenheit ? 32)" 
Set objRange = Selection.OMaths.Add(objRange) 
Set objEq = objRange.OMaths(1) 
objEq.BuildUp
```


## See also


[OMaths Object](Word.OMaths.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]