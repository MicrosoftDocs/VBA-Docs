---
title: CaptionLabels.Add method (Word)
keywords: vbawd10.chm158859364
f1_keywords:
- vbawd10.chm158859364
ms.prod: word
api_name:
- Word.CaptionLabels.Add
ms.assetid: f74af8c0-fa16-8ea2-3012-ac207d187502
ms.date: 06/08/2017
localization_priority: Normal
---


# CaptionLabels.Add method (Word)

Returns a  **CaptionLabel** object that represents a custom caption label.


## Syntax

_expression_.**Add** (_Name_)

_expression_ Required. A variable that represents a '[CaptionLabels](Word.captionlabels.md)' collection.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Name_|Required| **String**|The name of the custom caption label.|

## Return value

CaptionLabel


## Example

This example adds a custom caption label named Demo Slide. To verify that the custom label is added, view the **Label** combo box in the **Caption** dialog box, accessed from the **Reference** command on the **Insert** menu.


```vb
Sub CapLbl() 
 CaptionLabels.Add Name:="Demo Slide" 
End Sub
```


## See also


[CaptionLabels Collection Object](Word.captionlabels.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]