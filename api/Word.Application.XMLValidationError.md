---
title: Application.XMLValidationError event (Word)
keywords: vbawd10.chm4000026
f1_keywords:
- vbawd10.chm4000026
ms.prod: word
api_name:
- Word.Application.XMLValidationError
ms.assetid: bb75a555-fb5e-fb7b-f152-4c6436ecb1c7
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.XMLValidationError event (Word)

Occurs when there is a validation error in the document.


## Syntax

_expression_.**XMLValidationError'(**_XMLNode As XMLNode_**)

_expression_ A variable that represents an **[Application](Word.Application.md)** object.  An object of type **Application** that has been declared in a class module by using the **WithEvents** keyword. For more information about using events with the **Application** object, see [Using events with the Application object](../word/Concepts/Objects-Properties-Methods/using-events-with-the-application-object-word.md).


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _XMLNode_|Required| **XMLNode**|The XML element that is invalid.|

## Example

The following example displays an error message to the user when a node is invalid.


```vb
Private Sub Wrd_XMLValidationError(ByVal XMLNode As XMLNode) 
 MsgBox "The " & UCase(XMLNode.BaseName) & " element is invalid." & _ 
 vbCrLf & XMLNode.ValidationErrorText 
End Sub
```


## See also


[Application Object](Word.Application.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]