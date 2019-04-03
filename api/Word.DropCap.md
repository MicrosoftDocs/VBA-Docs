---
title: DropCap object (Word)
keywords: vbawd10.chm2390
f1_keywords:
- vbawd10.chm2390
ms.prod: word
api_name:
- Word.DropCap
ms.assetid: 79daea90-657b-43db-34e3-08f7aed74591
ms.date: 06/08/2017
localization_priority: Normal
---


# DropCap object (Word)

Represents a dropped capital letter at the beginning of a paragraph. There is no  **DropCaps** collection; each **[Paragraph](Word.Paragraph.md)** object contains only one **DropCap** object.


## Remarks

Use the  **DropCap** property to return a **DropCap** object. The following example sets a dropped capital letter for the first letter in the first paragraph in the active document.


```vb
With ActiveDocument.Paragraphs(1).DropCap 
 .Enable 
 .Position = wdDropNormal 
End With
```


## See also



[Word Object Model Reference](overview/Word/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]