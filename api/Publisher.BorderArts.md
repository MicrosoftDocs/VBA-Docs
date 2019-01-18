---
title: BorderArts Object (Publisher)
keywords: vbapb10.chm7798783
f1_keywords:
- vbapb10.chm7798783
ms.prod: publisher
api_name:
- Publisher.BorderArts
ms.assetid: 0fc016f6-154e-3591-34b3-e094bbad9d16
ms.date: 06/08/2017
localization_priority: Normal
---


# BorderArts Object (Publisher)

A collection of all BorderArt available for use in the specified publication. BorderArt is predefined picture borders that can be applied to text boxes, picture frames, or rectangles.
 


## Remarks

The  **BorderArts** collection includes any custom BorderArt types created by the user for the specified publication.
 

 

## Example

Use the  **[Item](Publisher.BorderArts.Item.md)** property of a **BorderArts** collection to return a specific **[BorderArt](Publisher.BorderArt.md)** object. The Index argument of the **Item** property can be the number or name of the BorderArt object.
 

 
This example returns the BorderArt "Apples" from the active publication. 
 

 



```vb
Dim bdaTemp As BorderArt 
 
Set bdaTemp = ActiveDocument.BorderArts.Item (Index:="Apples") 
```

Use the  **[Count](Publisher.BorderArts.Count.md)** property to return the number of BorderArt types available in the specified document. The following example displays the number of BorderArt types in the active document.
 

 



```vb
Sub CountBorderArts() 
 MsgBox ActiveDocument.BorderArts.Count 
End Sub
```


## Methods



|Name|
|:-----|
|[Item](Publisher.BorderArts.Item.md)|

## Properties



|Name|
|:-----|
|[Application](Publisher.BorderArts.Application.md)|
|[Count](Publisher.BorderArts.Count.md)|
|[Parent](Publisher.BorderArts.Parent.md)|

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]