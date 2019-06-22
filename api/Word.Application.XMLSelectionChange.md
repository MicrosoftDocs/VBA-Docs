---
title: Application.XMLSelectionChange event (Word)
keywords: vbawd10.chm4000025
f1_keywords:
- vbawd10.chm4000025
ms.prod: word
api_name:
- Word.Application.XMLSelectionChange
ms.assetid: a25d4e87-9b29-77b4-ddea-7692a0b56a8a
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.XMLSelectionChange event (Word)

Occurs when the parent XML node of the current selection changes.


## Syntax

_expression_.**XMLSelectionChange'(**_Sel_**, **_OldXMLNode_**, **_NewXMLNode_**, **_Reason_**)

_expression_ A variable that represents an '[Application](Word.Application.md)' object that has been declared in a class module by using the **WithEvents** keyword. For more information about using events with the **Application** object, see [Using events with the Application object](../word/Concepts/Objects-Properties-Methods/using-events-with-the-application-object-word.md).


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Sel_|Required| **Selection**|The text selected, including XML elements. If no text is selected, the Sel parameter returns either nothing or the first character to the right of the insertion point.|
| _OldXMLNode_|Required| **XMLNode**|The XML node from which the insertion point is moving.|
| _NewXMLNode_|Required| **XMLNode**|The XML node to which the insertion point is moving.|

## Example

The following example validates a newly added XML element when a new element is inserted into the document.


```vb
Private Sub Wrd_XMLSelectionChange(ByVal Sel As Selection, _ 
 ByVal OldXMLNode As XMLNode, ByVal NewXMLNode As XMLNode, _ 
 Reason As Long) 
 
 Dim intResponse As Integer 
 
 If Reason = wdXMLSelectionChangeReasonInsert Then 
 If Not NewXMLNode Is Nothing Then 
 NewXMLNode.Validate 
 End If 
 End If 
 
End Sub
```


## See also


[Application Object](Word.Application.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]