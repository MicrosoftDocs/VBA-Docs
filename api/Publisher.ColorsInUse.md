---
title: ColorsInUse Object (Publisher)
keywords: vbapb10.chm3014655
f1_keywords:
- vbapb10.chm3014655
ms.prod: publisher
api_name:
- Publisher.ColorsInUse
ms.assetid: ced0028a-8ab5-d9b1-b28c-24b794bdcbfe
ms.date: 06/08/2017
localization_priority: Normal
---


# ColorsInUse Object (Publisher)

A collection of  **[ColorFormat](Publisher.ColorFormat.md)** objects that represent the colors present in the specified publication.
 


## Remarks

The  **ColorsInUse** collection supports all the publication color models: RGB, process colors, and spot color.
 

 
For process color and spot color publications, colors are based on inks. For a given ink, a publication may contain several colors that are different tints or shades of that ink. Use the  **[Plates](Publisher.Plates.md)** collection to access the plates that represent the inks defined for a publication.
 

 

## Example

Use the  **[ColorsInUse](./overview/Publisher.md)** property of the **[Document](Publisher.Document.md)** object to return the **ColorsInUse** collection.
 

 
The following example lists properties of each color in the active publication that is based on the specified ink. This example assumes the publication's color mode has been defined as spot color or process and spot color.
 

 



```vb
Sub ListColorsBasedOnInk() 
Dim cfLoop As ColorFormat 
 
For Each cfLoop In ActiveDocument.ColorsInUse 
 
 With cfLoop 
 If .Ink = "2" Then 
 Debug.Print "BaseRGB: " &amp; .BaseRGB 
 Debug.Print "RGB: " &amp; .RGB 
 Debug.Print "TintShade: " &amp; .TintAndShade 
 Debug.Print "Type: " &amp; .Type 
 End If 
 End With 
 
Next cfLoop 
 
End Sub
```

Use  **ColorsInUse** (index), where index is the color index number, to return a single **ColorFormat** object. The following example returns properties for the second color in the publication.
 

 



```vb
Sub ColorProperties() 
 
 With ActiveDocument.ColorsInUse(2) 
 Debug.Print "Color RBG: " &amp; .RGB 
 Debug.Print "Ink RBG: " &amp; .BaseRGB 
 Debug.Print "Tint: " &amp; .TintAndShade 
 
 End With 
 
End Sub
```


## Properties



|Name|
|:-----|
|[Application](Publisher.ColorsInUse.Application.md)|
|[Count](Publisher.ColorsInUse.Count.md)|
|[Item](Publisher.ColorsInUse.Item.md)|
|[Parent](Publisher.ColorsInUse.Parent.md)|

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]