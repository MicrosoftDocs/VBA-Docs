---
title: Document.Variables property (Word)
keywords: vbawd10.chm158007322
f1_keywords:
- vbawd10.chm158007322
ms.prod: word
api_name:
- Word.Document.Variables
ms.assetid: 93af7b84-f172-6ebd-2147-e7ebc92449c5
ms.date: 06/08/2017
localization_priority: Normal
---


# Document.Variables property (Word)

Returns a  **[Variables](Word.variables.md)** collection that represents the variables stored in the specified document. Read-only.


## Syntax

_expression_. `Variables`

_expression_ A variable that represents a **[Document](Word.Document.md)** object.


## Remarks

For information about returning a single member of a collection, see [Returning an object from a collection](../word/Concepts/Miscellaneous/returning-an-object-from-a-collection-word.md).


## Example

This example adds a document variable named "Value1" to the active document. The example then retrieves the value from the Value1 variable, adds 3 to the value, and displays the results.


```vb
ActiveDocument.Variables.Add Name:="Value1", Value:="1" 
MsgBox ActiveDocument.Variables("Value1") + 3
```

This example displays the name and value of each document variable in the active document.




```vb
For Each myVar In ActiveDocument.Variables 
 MsgBox "Name =" & myVar.Name & vbCr & "Value = " & myVar.Value 
Next myVar
```


## See also


[Document Object](Word.Document.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]