---
title: ScratchArea object (Publisher)
keywords: vbapb10.chm1245183
f1_keywords:
- vbapb10.chm1245183
api_name:
- Publisher.ScratchArea
ms.assetid: 41856866-c1d8-2550-1b4c-28886ed2b714
ms.date: 06/01/2019
ms.localizationpriority: medium
---


# ScratchArea object (Publisher)

Represents the area outside the boundaries of publication pages where layout elements may be stored with no effect on publication output.
 
## Remarks

Use the **[ScratchArea](Publisher.Document.ScratchArea.md)** property of the **Document** object to return a scratch area. 

Use the **Shapes** property to return the collection of shapes that are currently on a scratch area.

## Example
 
This example assigns the first shape on the scratch area of the active document to a variable.
 
```vb
Dim saPage As ScratchArea 
Dim objFirst As Object 
 
saPage = Application.ActiveDocument.ScratchArea 
objFirst = saPage.Shapes(1)
```


## Properties

- [Application](Publisher.ScratchArea.Application.md)
- [Parent](Publisher.ScratchArea.Parent.md)
- [Shapes](Publisher.ScratchArea.Shapes.md)


## See also

- [Publisher Object Model Reference](overview/publisher/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]