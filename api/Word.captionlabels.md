---
title: CaptionLabels Object (Word)
ms.prod: word
ms.assetid: 7d18c0d6-6d58-9841-4665-ab13e2e2ad9f
ms.date: 06/08/2017
---


# CaptionLabels Object (Word)

A collection of  **[CaptionLabel](Word.CaptionLabel.md)** objects that represent the available caption labels. The items in the **CaptionLabels** collection are listed in the **Label** box in the **Caption** dialog box.


## Remarks

Use the  **CaptionLabels** property to return the **CaptionLabels** collection. By default, the **CaptionLabels** collection includes the three built-in caption labels: Figure, Table, and Equation.

Use the  **[Add](Word.CaptionLabels.Add.md)** method to add a custom caption label. The following example adds a caption label named "Photo."




```vb
CaptionLabels.Add Name:="Photo"
```

Use  **CaptionLabels** (index), where index is the caption label name or index number, to return a single **CaptionLabel** object. The following example sets the numbering style for the Figure caption label.




```vb
CaptionLabels("Figure").NumberStyle = _ 
 wdCaptionNumberStyleLowercaseLetter
```

The index number represents the position of the caption label in the  **CaptionLabels** collection. The following example displays the first caption label.




```vb
MsgBox CaptionLabels(1).Name
```


## Methods



|**Name**|
|:-----|
|[Add](Word.CaptionLabels.Add.md)|
|[Item](Word.CaptionLabels.Item.md)|

## Properties



|**Name**|
|:-----|
|[Application](Word.CaptionLabels.Application.md)|
|[Count](Word.CaptionLabels.Count.md)|
|[Creator](Word.CaptionLabels.Creator.md)|
|[Parent](captionlabels-parent-property-word.md)|

## See also


[Word Object Model Reference](./overview/object-model-word-vba-reference.md)
