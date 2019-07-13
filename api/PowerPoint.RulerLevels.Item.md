---
title: RulerLevels.Item method (PowerPoint)
keywords: vbapp10.chm571003
f1_keywords:
- vbapp10.chm571003
ms.prod: powerpoint
api_name:
- PowerPoint.RulerLevels.Item
ms.assetid: 95c04d29-0c1c-9df0-6d6d-43da01ea7ae2
ms.date: 06/08/2017
localization_priority: Normal
---


# RulerLevels.Item method (PowerPoint)

Returns a single  **RulerLevel** object from the specified **RulerLevels** collection.


## Syntax

_expression_.**Item** (_Index_)

_expression_ A variable that represents a [RulerLevels](PowerPoint.RulerLevels.md) collection.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Index_|Required|**Long**|The index number of the single  **RulerLevel** object in the collection to be returned.|

## Return value

RulerLevel


## Example

This example sets the first-line indent and the hanging indent for outline level one in body text on the slide master for the active presentation.


```vb
With ActivePresentation.SlideMaster.TextStyles.Item(ppBodyStyle)

    With .Ruler.Levels.Item(1) ' sets indents for level 1

        .FirstMargin = 9

        .LeftMargin = 54

    End With

End With
```


## See also


[RulerLevels Object](PowerPoint.RulerLevels.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]