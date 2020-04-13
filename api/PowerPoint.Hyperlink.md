---
title: Hyperlink object (PowerPoint)
keywords: vbapp10.chm526000
f1_keywords:
- vbapp10.chm526000
ms.prod: powerpoint
api_name:
- PowerPoint.Hyperlink
ms.assetid: c8d53079-b280-c93c-a3c9-b865d09abe1a
ms.date: 06/08/2017
localization_priority: Normal
---


# Hyperlink object (PowerPoint)

Represents a hyperlink associated with a non-placeholder shape or text. 


## Remarks

You can use a hyperlink to jump to an Internet or intranet site, to another file, or to a slide within the active presentation. The **Hyperlink** object is a member of the **[Hyperlinks](PowerPoint.Hyperlinks.md)** collection. The **Hyperlinks** collection contains all the hyperlinks on a slide or a master.


## Example

Use the [Hyperlink](PowerPoint.ActionSetting.Hyperlink.md)property to return a hyperlink for a shape. A shape can have two different hyperlinks assigned to it: one that is followed when the user clicks the shape during a slide show, and another that is followed when the user passes the mouse pointer over the shape during a slide show. For the hyperlink to be active during a slide show, the  **Action** property must be set to **ppActionHyperlink**. The following example sets the mouse-click action for shape three on slide one in the active presentation to an Internet link.


```vb
With ActivePresentation.Slides(1).Shapes(3) _

        .ActionSettings(ppMouseClick)

    .Action = ppActionHyperlink

    .Hyperlink.Address = "https://www.microsoft.com"

End With
```

A slide can contain more than one hyperlink. Each non-placeholder shape can have a hyperlink; the text within a shape can have its own hyperlink; and each individual character can have its own hyperlink. Use  **Hyperlinks** (_index_), where _index_ is the hyperlink number, to return a single **Hyperlink** object. The following example adds the shape three mouse-click hyperlink to the Favorites folder.




```vb
ActivePresentation.Slides(1).Shapes(3) _

    .ActionSettings(ppMouseClick).Hyperlink.AddToFavorites
```


> [!NOTE] 
> When you use this method to add a hyperlink to the Internet Explorer Favorites folder, an icon is added to the  **Favorites** menu without a corresponding name. You must add the name from within Internet Explorer.


## Methods



|Name|
|:-----|
|[AddToFavorites](PowerPoint.Hyperlink.AddToFavorites.md)|
|[CreateNewDocument](PowerPoint.Hyperlink.CreateNewDocument.md)|
|[Delete](PowerPoint.Hyperlink.Delete.md)|
|[Follow](PowerPoint.Hyperlink.Follow.md)|

## Properties



|Name|
|:-----|
|[Address](PowerPoint.Hyperlink.Address.md)|
|[Application](PowerPoint.Hyperlink.Application.md)|
|[EmailSubject](PowerPoint.Hyperlink.EmailSubject.md)|
|[Parent](PowerPoint.Hyperlink.Parent.md)|
|[ScreenTip](PowerPoint.Hyperlink.ScreenTip.md)|
|[ShowAndReturn](PowerPoint.Hyperlink.ShowAndReturn.md)|
|[SubAddress](PowerPoint.Hyperlink.SubAddress.md)|
|[TextToDisplay](PowerPoint.Hyperlink.TextToDisplay.md)|
|[Type](PowerPoint.Hyperlink.Type.md)|

## See also


[PowerPoint Object Model Reference](overview/PowerPoint/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
