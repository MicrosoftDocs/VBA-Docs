---
title: Presentation.Merge method (PowerPoint)
keywords: vbapp10.chm583064
f1_keywords:
- vbapp10.chm583064
ms.assetid: 5cc604de-6d57-69dc-e3bc-88505b947f72
ms.date: 06/08/2017
ms.localizationpriority: medium
---


# Presentation.Merge method (PowerPoint)
> [!NOTE] 
> This object or member has been deprecated, but it remains part of the object model for backward compatibility. You should not use it in new applications.

Merges the changes in one presentation with another.


## Syntax

_expression_.**Merge** (_Path_)

_expression_ A variable that represents a [Presentation](PowerPoint.Presentation.md) object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Path_|Required|**String**|The path, including filename, of the presentation to merge changes with.|


## Return value

 **VOID**


## Example

The following code sample merges the active presentation with a presentation saved to the user's desktop.


```vb
Sub MergePresentations()
    Dim userName As String
    Dim otherPres As String

    ActivePresentation.Merge("C:\Users\? & username & ?\Desktop\" & otherPres)
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
