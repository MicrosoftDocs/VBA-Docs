---
title: XMLNode object (Word)
keywords: vbawd10.chm576
f1_keywords:
- vbawd10.chm576
ms.prod: word
api_name:
- Word.XMLNode
ms.assetid: fe305ba9-7375-ad4f-6036-155add17a9d0
ms.date: 06/08/2017
localization_priority: Normal
---


# XMLNode object (Word)

Represents a single XML element applied to a document. 


## Remarks

Each XML element that has been applied to a document is displayed as a node in a tree view control in the **XML Structure** task pane. Each node in the tree view is an instance of an **XMLNode** object. The hierarchy in the tree view indicates whether a node contains child nodes.

Use the **Item** method of the **XMLNodes** collection to return an individual **XMLNode** object. Use the **Validate** method to verify that an XML element is valid according to the applied schemas and that any required child elements exist and are in the required order. Once you run the **Validate** method, use the **ValidationStatus** property to verify whether an element is valid, and use the **ValidationErrorText** property to display information about what the user needs to do to make the document conform to the XML schema rules.

The following example validates each of the XML elements in the active document. If the element is found to be invalid against the schema, the example returns a message to the user explaining what the problem is.




```vb
Sub ValidateXMLElements() 
 Dim objNode As XMLNode 
 
 For Each objNode In ActiveDocument.XMLNodes 
 objNode.Validate 
 If objNode.ValidationStatus <> wdXMLValidationStatusOK Then 
 MsgBox objNode.ValidationErrorText(True) 
 End If 
 Next 
End Sub
```


## See also


[Word Object Model Reference](overview/Word/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]