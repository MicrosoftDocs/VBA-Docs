---
title: ListLevel object (Word)
keywords: vbawd10.chm2445
f1_keywords:
- vbawd10.chm2445
ms.prod: word
api_name:
- Word.ListLevel
ms.assetid: 0cd152cb-6c25-50cb-7c1d-8b6d9734505b
ms.date: 06/08/2017
localization_priority: Normal
---


# ListLevel object (Word)

Represents a single list level, either the only level for a bulleted or numbered list or one of the nine levels of an outline numbered list. The  **ListLevel** object is a member of the **ListLevels** collection.


## Remarks

Use  **ListLevels** (Index), where Index is a number from 1 through 9, to return a single **ListLevel** object. The following example sets list level one of list template one in the active document to start at 4.


```vb
ActiveDocument.ListTemplates(1).ListLevels(1).StartAt = 4
```

The  **ListLevel** object gives you access to all the formatting properties for the specified list level, such as the **Alignment**, **Font**, **NumberFormat**, **NumberPosition**, **NumberStyle**, and **TrailingCharacter** properties.

To apply a list level, first identify the range or list, and then use the  **ApplyListTemplate** method. Each tab at the beginning of the paragraph is translated into a list level. For example, a paragraph that begins with three tabs will become a level-three list paragraph after the **ApplyListTemplate** method is used.


## Methods



|Name|
|:-----|
|[ApplyPictureBullet](Word.ListLevel.ApplyPictureBullet.md)|

## Properties



|Name|
|:-----|
|[Alignment](Word.ListLevel.Alignment.md)|
|[Application](Word.ListLevel.Application.md)|
|[Creator](Word.ListLevel.Creator.md)|
|[Font](Word.ListLevel.Font.md)|
|[Index](Word.ListLevel.Index.md)|
|[LinkedStyle](Word.ListLevel.LinkedStyle.md)|
|[NumberFormat](Word.ListLevel.NumberFormat.md)|
|[NumberPosition](Word.ListLevel.NumberPosition.md)|
|[NumberStyle](Word.ListLevel.NumberStyle.md)|
|[Parent](Word.ListLevel.Parent.md)|
|[PictureBullet](Word.ListLevel.PictureBullet.md)|
|[ResetOnHigher](Word.ListLevel.ResetOnHigher.md)|
|[StartAt](Word.ListLevel.StartAt.md)|
|[TabPosition](Word.ListLevel.TabPosition.md)|
|[TextPosition](Word.ListLevel.TextPosition.md)|
|[TrailingCharacter](Word.ListLevel.TrailingCharacter.md)|

## See also


[Word Object Model Reference](overview/Word/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]