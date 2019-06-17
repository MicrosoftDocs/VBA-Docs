---
title: WebPageOptions.Description property (Publisher)
keywords: vbapb10.chm544771
f1_keywords:
- vbapb10.chm544771
ms.prod: publisher
api_name:
- Publisher.WebPageOptions.Description
ms.assetid: dfd18427-c70d-7232-191e-a6332a89c3fe
ms.date: 06/18/2019
localization_priority: Normal
---


# WebPageOptions.Description property (Publisher)

Returns or sets a **String** that represents the description of a webpage within a web publication. Read/write.


## Syntax

_expression_.**Description**

_expression_ A variable that represents a **[WebPageOptions](Publisher.WebPageOptions.md)** object.


## Example

This example sets the description for page two of the active web publication.

```vb
Dim theWPO As WebPageOptions 
 
Set theWPO = ActiveDocument.Pages(2).WebPageOptions 
 
With theWPO 
 .Description = "Company Profile" 
End With
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]