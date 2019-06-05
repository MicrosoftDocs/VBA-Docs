---
title: Document.BorderArts property (Publisher)
keywords: vbapb10.chm196721
f1_keywords:
- vbapb10.chm196721
ms.prod: publisher
api_name:
- Publisher.Document.BorderArts
ms.assetid: 5639ffce-f711-71b6-78f8-2de63fe50a3c
ms.date: 06/06/2019
localization_priority: Normal
---


# Document.BorderArts property (Publisher)

Returns a **[BorderArts](Publisher.BorderArts.md)** collection that represents the BorderArt types available for use in the specified publication. Read-only.


## Syntax

_expression_.**BorderArts**

_expression_ A variable that represents a **[Document](Publisher.Document.md)** object.


## Return value

BorderArts


## Remarks

BorderArt are picture borders that can be applied to text boxes, picture frames, or rectangles. 

The **BorderArts** collection includes any custom BorderArt types created by the user for the specified publication.


## Example

The following example returns the **BorderArts** collection and lists the names of all the BorderArt types available for use in the active publication.

```vb
Sub ListBorderArt() 
Dim bdaTemp As BorderArts 
Dim bdaLoop As BorderArt 
 
Set bdaTemp = ActiveDocument.BorderArts 
 
For Each bdaLoop In bdaTemp 
 Debug.Print "The name of this BorderArt is " & bdaLoop.Name 
Next bdaLoop 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]