---
title: Presentation.Merge Method (PowerPoint)
keywords: vbapp10.chm583064
f1_keywords:
- vbapp10.chm583064
ms.assetid: 5cc604de-6d57-69dc-e3bc-88505b947f72
ms.date: 06/08/2017
ms.prod: powerpoint
localization_priority: Normal
---


# Presentation.Merge Method (PowerPoint)

Merges the changes in one presentation with another.


## Syntax

 _expression_. `Merge`_(Path)_

 _expression_ A variable that represents a [Presentation](./PowerPoint.Presentation.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Path_|Required|**String**|The path, including filename, of the presentation to merge changes with.|
| _Path_|Required|STRING||

## Return value

 **VOID**


## Example

The following code sample merges the active presentation with a presentation saved to the user?s desktop.


```vb
Sub MergePresentations()
    Dim userName As String
    Dim otherPres As String

    ActivePresentation.Merge("C:\Users\? & username & ?\Desktop\" & otherPres)
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]