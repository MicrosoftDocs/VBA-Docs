---
title: CaptionLabel object (Word)
keywords: vbawd10.chm2425
f1_keywords:
- vbawd10.chm2425
ms.prod: word
api_name:
- Word.CaptionLabel
ms.assetid: 71c82dfd-6a66-e0f4-e30f-ae453c764864
ms.date: 06/08/2017
localization_priority: Normal
---


# CaptionLabel object (Word)

Represents a single caption label. The **CaptionLabel** object is a member of the **[CaptionLabels](Word.captionlabels.md)** collection. The items in the **CaptionLabels** collection are listed in the **Label** box in the **Caption** dialog box.


## Remarks

Use  **[CaptionLabels](Word.Application.CaptionLabels.md)** (_index_), where _index_ is the caption label name or index number, to return a single **CaptionLabel** object. The following example sets the numbering style for the Figure caption label.


```vb
CaptionLabels("Figure").NumberStyle = _ 
 wdCaptionNumberStyleLowercaseLetter
```

The index number represents the position of the caption label in the **CaptionLabels** collection. The following example displays the first caption label.




```vb
MsgBox CaptionLabels(1).Name
```

Use the **[Add](Word.CaptionLabels.Add.md)** method to add a custom caption label. The following example adds a caption label named "Photo."




```vb
CaptionLabels.Add Name:="Photo"
```


## Methods



|Name|
|:-----|
|[Delete](Word.CaptionLabel.Delete.md)|

## Properties



|Name|
|:-----|
|[Application](Word.CaptionLabel.Application.md)|
|[BuiltIn](Word.CaptionLabel.BuiltIn.md)|
|[ChapterStyleLevel](Word.CaptionLabel.ChapterStyleLevel.md)|
|[Creator](Word.CaptionLabel.Creator.md)|
|[ID](Word.CaptionLabel.ID.md)|
|[IncludeChapterNumber](Word.CaptionLabel.IncludeChapterNumber.md)|
|[Name](Word.CaptionLabel.Name.md)|
|[NumberStyle](Word.CaptionLabel.NumberStyle.md)|
|[Parent](Word.CaptionLabel.Parent.md)|
|[Position](Word.CaptionLabel.Position.md)|
|[Separator](Word.CaptionLabel.Separator.md)|

## See also


[Word Object Model Reference](overview/Word/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]