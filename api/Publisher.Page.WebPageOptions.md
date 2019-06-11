---
title: Page.WebPageOptions property (Publisher)
keywords: vbapb10.chm393264
f1_keywords:
- vbapb10.chm393264
ms.prod: publisher
api_name:
- Publisher.Page.WebPageOptions
ms.assetid: c2e3ee01-5b49-e83c-a68b-a4d526da0215
ms.date: 06/11/2019
localization_priority: Normal
---


# Page.WebPageOptions property (Publisher)

Returns a **[WebPageOptions](Publisher.WebPageOptions.md)** object, which represents the properties of a single webpage within a web publication. Read-only.


## Syntax

_expression_.**WebPageOptions**

_expression_ A variable that represents a **[Page](Publisher.Page.md)** object.


## Return value

WebPageOptions


## Example

The following example sets the description and the background sound for the fourth page of the active web publication.

```vb
With ActiveDocument.Pages(4).WebPageOptions 
 .Description = "Company Profile" 
 .BackgroundSound = "C:\CompanySounds\corporate_jingle.wav" 
End With 

```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]