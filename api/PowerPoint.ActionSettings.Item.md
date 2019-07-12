---
title: ActionSettings.Item method (PowerPoint)
keywords: vbapp10.chm566003
f1_keywords:
- vbapp10.chm566003
ms.prod: powerpoint
api_name:
- PowerPoint.ActionSettings.Item
ms.assetid: 88e0b49b-0518-559b-243f-c369c09ab3fe
ms.date: 06/08/2017
localization_priority: Normal
---


# ActionSettings.Item method (PowerPoint)

Returns a single action setting from the specified **ActionSettings** collection.


## Syntax

_expression_.**Item** (_Index_)

_expression_ A variable that represents an **[ActionSettings](PowerPoint.ActionSettings.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Index_|Required|**PpMouseActivation**|The action setting for a **MouseClick** or **MouseOver** event.|

## Return value

ActionSetting


## Remarks

The _Index_ parameter value can be one of these **PpMouseActivation** constants.

|Constant|Description|
|:-----|:-----|
|**ppMouseClick**|The action setting for when the user clicks the shape.|
|**ppMouseOver**|The action setting for when the mouse pointer is positioned over the specified shape.|

The **Item** method is the default member for a collection. For example, the following two lines of code are equivalent:

```vb
ActivePresentation.Slides.Item(1)
```

```vb
ActivePresentation.Slides(1)
```

For more information about returning a single member of a collection, see [Returning an object from a collection](../powerpoint/How-to/return-objects-from-collections.md).


## Example

This example sets shape three on slide one to play the sound of applause and uses the **[AnimateAction](PowerPoint.ActionSetting.AnimateAction.md)** property to specify that the shape's color is to be momentarily inverted when the shape is clicked during a slide show.


```vb
With ActivePresentation.Slides.Item(1).Shapes _
        .Item(3).ActionSettings.Item(ppMouseClick)
    .SoundEffect.Name = "applause"
    .AnimateAction = True
End With
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]