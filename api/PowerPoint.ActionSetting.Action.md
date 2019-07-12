---
title: ActionSetting.Action property (PowerPoint)
keywords: vbapp10.chm567003
f1_keywords:
- vbapp10.chm567003
ms.prod: powerpoint
api_name:
- PowerPoint.ActionSetting.Action
ms.assetid: 32ed5574-5ac0-abb7-d300-6644fc894ec1
ms.date: 06/08/2017
localization_priority: Normal
---


# ActionSetting.Action property (PowerPoint)

Returns or sets the type of action that will occur when the specified shape is clicked or the mouse pointer is positioned over the shape during a slide show. Read/write.


## Syntax

_expression_.**Action**

_expression_ A variable that represents an **[ActionSetting](PowerPoint.ActionSetting.md)** object.


## Return value

Long


## Remarks

The **Action** property value can be one of the **[PpActionType](powerpoint.ppactiontype.md)** constants.

You can use the **Action** property in conjunction with other properties of the **ActionSetting** object, as shown in the following table.

|If you set the Action property to this value|Use this property|To do this|
|:-----|:-----|:-----|
|**ppActionHyperlink**|[Hyperlink](PowerPoint.ActionSetting.Hyperlink.md)|Set properties for the hyperlink that will be followed in response to a mouse action on the shape during a slide show.|
|**ppActionRunProgram**|**[Run](PowerPoint.ActionSetting.Run.md)**|Return or set the name of the program to run in response to a mouse action on the shape during a slide show.|
|**ppActionRunMacro**|**[Run](PowerPoint.ActionSetting.Run.md)**|Return or set the name of the macro to run in response to a mouse action on the shape during a slide show.|
|**ppActionOLEVerb**|[ActionVerb](PowerPoint.ActionSetting.ActionVerb.md)|Set the OLE verb that will be invoked in response to a mouse action on the shape during a slide show.|
|**ppActionNamedSlideShow**|[SlideShowName](PowerPoint.ActionSetting.SlideShowName.md)|Set the name of the custom slide show that will run in response to a mouse action on the shape during a slide show.|

## Example

This example sets shape three (an OLE object) on slide one in the active presentation to be played when the mouse passes over it during a slide show.


```vb
With ActivePresentation.Slides(1) _
    .Shapes(3).ActionSettings(ppMouseOver)

        .ActionVerb = "Play"
        .Action = ppActionOLEVerb

End With
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]