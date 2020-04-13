---
title: ActionSetting object (PowerPoint)
keywords: vbapp10.chm567000
f1_keywords:
- vbapp10.chm567000
ms.prod: powerpoint
api_name:
- PowerPoint.ActionSetting
ms.assetid: 21381ff0-b9ff-59d8-77e9-345905fb8617
ms.date: 06/08/2017
localization_priority: Normal
---


# ActionSetting object (PowerPoint)

Contains information about how the specified shape or text range reacts to mouse actions during a slide show. 


## Remarks

The **ActionSetting** object is a member of the **[ActionSettings](PowerPoint.ActionSettings.md)** collection. The **ActionSettings** collection contains one **ActionSetting** object that represents how the specified object reacts when the user clicks it during a slide show and one **ActionSetting** object that represents how the specified object reacts when the user moves the mouse pointer over it during a slide show.

If you've set properties of the  **ActionSetting** object that don't seem to be taking effect, make sure that you've set the [Action](PowerPoint.ActionSetting.Action.md) property to the appropriate value.


## Example

Use  **ActionSettings** (_index_), where _index_ is the either **ppMouseClick** or **ppMouseOver**, to return a single **ActionSetting** object. The following example sets the mouse-click action for the text in the third shape on slide one in the active presentation to an Internet link.


```vb
With ActivePresentation.Slides(1).Shapes(3) _ 
        .TextFrame.TextRange.ActionSettings(ppMouseClick) 
    .Action = ppActionHyperlink 
    .Hyperlink.Address = "https://www.microsoft.com" 
End With
```


## Properties



|Name|
|:-----|
|[Action](PowerPoint.ActionSetting.Action.md)|
|[ActionVerb](PowerPoint.ActionSetting.ActionVerb.md)|
|[AnimateAction](PowerPoint.ActionSetting.AnimateAction.md)|
|[Application](PowerPoint.ActionSetting.Application.md)|
|[Hyperlink](PowerPoint.ActionSetting.Hyperlink.md)|
|[Parent](PowerPoint.ActionSetting.Parent.md)|
|[Run](PowerPoint.ActionSetting.Run.md)|
|[ShowAndReturn](PowerPoint.ActionSetting.ShowAndReturn.md)|
|[SlideShowName](PowerPoint.ActionSetting.SlideShowName.md)|
|[SoundEffect](PowerPoint.ActionSetting.SoundEffect.md)|

## See also


[PowerPoint Object Model Reference](overview/PowerPoint/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
