---
title: ColorSchemes object (Publisher)
keywords: vbapb10.chm2818047
f1_keywords:
- vbapb10.chm2818047
ms.prod: publisher
api_name:
- Publisher.ColorSchemes
ms.assetid: f5002de1-5e91-fc92-eedb-0e13dce57802
ms.date: 05/31/2019
localization_priority: Normal
---


# ColorSchemes object (Publisher)

A collection of all the **[ColorScheme](Publisher.ColorScheme.md)** objects in Microsoft Publisher. Each **ColorScheme** object represents a color scheme, which is a set of colors that are used in a publication.
 
## Remarks

Use the **Count** property to return the number of color schemes available to Publisher. 

Use the **Item** property to return a specific color scheme from the **ColorSchemes** collection. The _Index_ argument of the **Item** property can be the number or name of the color scheme or a **[PbColorScheme](publisher.pbcolorscheme.md)** constant. 

Use the **[Name](Publisher.ColorScheme.Name.md)** property to return a color scheme name. 

## Example

The following example displays the number of color schemes.

```vb
Sub CountColorSchemes() 
 MsgBox Application.ColorSchemes.Count 
End Sub
```

<br/>

The follow example sets the color scheme of the active publication to Wildflower.

```vb
Sub SetColorScheme() 
 ActiveDocument.ColorScheme _ 
 = ColorSchemes.Item(pbColorSchemeWildflower) 
End Sub
```

<br/>

The following example lists in a text box all the color schemes available to Publisher.

```vb
Sub ListColorShemes() 
 
 Dim clrScheme As ColorScheme 
 Dim strSchemes As String 
 
 For Each clrScheme In Application.ColorSchemes 
 strSchemes = strSchemes & clrScheme.Name & vbLf 
 Next 
 ActiveDocument.Pages(1).Shapes.AddTextbox( _ 
 Orientation:=pbTextOrientationHorizontal, _ 
 Left:=72, Top:=72, Width:=400, Height:=500).TextFrame _ 
 .TextRange.Text = strSchemes 
 
End Sub
```


## Properties

- [Application](Publisher.ColorSchemes.Application.md)
- [Count](Publisher.ColorSchemes.Count.md)
- [Item](Publisher.ColorSchemes.Item.md)
- [Parent](Publisher.ColorSchemes.Parent.md)

## See also

- [Publisher Object Model Reference](overview/publisher/object-model.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]