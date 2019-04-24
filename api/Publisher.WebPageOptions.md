---
title: WebPageOptions object (Publisher)
keywords: vbapb10.chm548863
f1_keywords:
- vbapb10.chm548863
ms.prod: publisher
api_name:
- Publisher.WebPageOptions
ms.assetid: 694b56ce-1c2d-8202-25b7-19e55aadb0fd
ms.date: 06/08/2017
localization_priority: Normal
---


# WebPageOptions object (Publisher)

Represents the properties of a single Web page within a web publication, including options for adding the title and description of the page, background sounds, in addition to other options. The  **WebPageOptions** object is a member of the **[Page](Publisher.Page.md)** object.
 


## Remarks

Note that the  **WebPageOptions** object is only available when the active publication is a web publication. A run-time error is returned if trying to access this object from a print publication.
 

 

## Example

Use the  **[WebPageOptions](Publisher.Page.WebPageOptions.md)** property on the **Page** object to return a **WebPageOptions** object. Use the **[Description](Publisher.WebPageOptions.Description.md)** property to set the description of a specified Web page. The following example sets the description for the second page of the active web publication.
 

 

```vb
Dim theWPO As WebPageOptions 
 
Set theWPO = ActiveDocument.Pages(2).WebPageOptions 
 
With theWPO 
 .Description = "Company Profile" 
End With
```


## Methods



|Name|
|:-----|
|[SetBackgroundSoundRepeat](Publisher.WebPageOptions.SetBackgroundSoundRepeat.md)|

## Properties



|Name|
|:-----|
|[Application](Publisher.WebPageOptions.Application.md)|
|[BackgroundSound](Publisher.WebPageOptions.BackgroundSound.md)|
|[BackgroundSoundLoopCount](Publisher.WebPageOptions.BackgroundSoundLoopCount.md)|
|[BackgroundSoundLoopForever](Publisher.WebPageOptions.BackgroundSoundLoopForever.md)|
|[Description](Publisher.WebPageOptions.Description.md)|
|[IncludePageOnNewWebNavigationBars](Publisher.WebPageOptions.IncludePageOnNewWebNavigationBars.md)|
|[Keywords](Publisher.WebPageOptions.Keywords.md)|
|[Parent](Publisher.WebPageOptions.Parent.md)|
|[PublishFileName](Publisher.WebPageOptions.PublishFileName.md)|

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]