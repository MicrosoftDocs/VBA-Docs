---
title: Document.CompatibilityMode property (Word)
keywords: vbawd10.chm158007863
f1_keywords:
- vbawd10.chm158007863
ms.prod: word
api_name:
- Word.Document.CompatibilityMode
ms.assetid: 5e4be325-1883-7701-53a1-4d7e20e3a989
ms.date: 06/08/2017
localization_priority: Normal
---


# Document.CompatibilityMode property (Word)

Returns a  **Long** that specifies the compatibility mode that Word uses when opening the document. Read-only.


## Syntax

 _expression_. `CompatibilityMode`

 _expression_ An expression that returns a '[Document](Word.Document.md)' object.


## Remarks

When you open a document in Word that was created in a previous version of Word, Compatibility Mode is turned on. Compatibility Mode ensures that no new or enhanced features in Word are available while working with a document, so that people who edit the document using previous versions of Word will have full editing capabilities.


## Example

The following example shows how to check if a document is in full fidelity mode before using a new feature. In this case, if the document compatibility mode supports using content controls, then a check box content control is added to the document.


```vb
Sub InsertCheckbox()
       
    If (Application.Version = ActiveDocument.CompatibilityMode) Then
          Selection.Range.ContentControls.Add (wdContentControlCheckBox)
    End If    
End Sub
```


## See also


[Document Object](Word.Document.md)

