---
title: Selection.PreviousRevision method (Word)
keywords: vbawd10.chm158663188
f1_keywords:
- vbawd10.chm158663188
ms.prod: word
api_name:
- Word.Selection.PreviousRevision
ms.assetid: e516037f-047d-5cd2-19b4-3b7870a14b5a
ms.date: 06/08/2017
localization_priority: Normal
---


# Selection.PreviousRevision method (Word)

Locates and returns the previous tracked change as a  **Revision** object.


## Syntax

_expression_. `PreviousRevision`( `_Wrap_` )

_expression_ Required. A variable that represents a **[Selection](Word.Selection.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Wrap_|Optional| **Variant**| **True** to continue searching for a revision at the end of the document when the beginning of the document is reached. The default value is **False**.|

## Return value

Revision


## Example

This example selects the last tracked change in the first section in the active document and displays the date and time of the change.


```vb
Selection.EndOf Unit:=wdStory, Extend:=wdMove 
Set myRev = Selection.PreviousRevision 
If Not (myRev Is Nothing) Then MsgBox myRev.Date
```

This example rejects the previous tracked change found if the change type is deleted or inserted text. If the tracked change is a style change, the change is accepted.




```vb
Set myRev = Selection.PreviousRevision(Wrap:=True) 
If Not (myRev Is Nothing) Then 
 Select Case myRev.Type 
 Case wdRevisionDelete 
 myRev.Reject 
 Case wdRevisionInsert 
 myRev.Reject 
 Case wdRevisionStyle 
 myRev.Accept 
 End Select 
End If
```


## See also


[Selection Object](Word.Selection.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]