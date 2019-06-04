---
title: BorderArtFormat.Set method (Publisher)
keywords: vbapb10.chm7602185
f1_keywords:
- vbapb10.chm7602185
ms.prod: publisher
api_name:
- Publisher.BorderArtFormat.Set
ms.assetid: e068037b-56b6-a114-6b22-568ea20d6b25
ms.date: 06/05/2019
localization_priority: Normal
---


# BorderArtFormat.Set method (Publisher)

Sets the type of BorderArt applied to the specified shape.


## Syntax

_expression_.**Set** (_BorderArtName_)

_expression_ A variable that represents a **[BorderArtFormat](Publisher.BorderArtFormat.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_BorderArtName_|Required| **Variant**|The name of the BorderArt type applied to the specified shape.|

## Remarks

You can also set the type of BorderArt applied to a shape by using the **[Name](Publisher.BorderArtFormat.Name.md)** property.


## Example

The following example tests for the existence of BorderArt on each shape for each page of the active document. Any BorderArt found is set to the same type.

```vb
Sub SetBorderArt() 
Dim anyPage As Page 
Dim anyShape As Shape 
Dim strBorderArtName As String 
 
strBorderArtName = Document.BorderArts(1).Name 
 
For Each anyPage in ActiveDocument.Pages 
 For Each anyShape in anyPage.Shapes 
 With anyShape.BorderArt 
 If .Exists = True Then 
 .Set(strBorderArtName) 
 End If 
 End With 
 Next anyShape 
 Next anyPage 
End Sub
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]