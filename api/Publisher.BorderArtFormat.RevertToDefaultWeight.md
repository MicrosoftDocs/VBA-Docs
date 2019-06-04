---
title: BorderArtFormat.RevertToDefaultWeight method (Publisher)
keywords: vbapb10.chm7602180
f1_keywords:
- vbapb10.chm7602180
ms.prod: publisher
api_name:
- Publisher.BorderArtFormat.RevertToDefaultWeight
ms.assetid: 3e46637f-3fce-3346-9193-063be40844bd
ms.date: 06/05/2019
localization_priority: Normal
---


# BorderArtFormat.RevertToDefaultWeight method (Publisher)

Sets the BorderArt on the specified shape back to its default thickness.


## Syntax

_expression_.**RevertToDefaultWeight**

_expression_ A variable that represents a **[BorderArtFormat](Publisher.BorderArtFormat.md)** object.


## Remarks

The **RevertToDefaultWeight** method has the same effect as the **Always apply at default size** control in the **BorderArt** dialog box.

Use the **[Weight](Publisher.BorderArtFormat.Weight.md)** property to set the specified BorderArt to a thickness other than the default.


## Example

The following example tests for the existence of BorderArt on each shape for each page of the active document. If BorderArt exists, its weight is set to the default thickness and original color.

```vb
Sub RestoreBorderArtDefaults() 
 
Dim anyPage As Page 
Dim anyShape As Shape 
 
For Each anyPage in ActiveDocument.Pages 
 For Each anyShape in anyPage.Shapes 
 With anyShape.BorderArt 
 If .Exists = True Then 
 .RevertToDefaultWeight 
 .RevertToOriginalColor 
 End If 
 End With 
 Next anyShape 
Next anyPage 
End Sub
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]