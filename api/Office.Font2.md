---
title: Font2 Object (Office)
ms.prod: office
api_name:
- Office.Font2
ms.assetid: 8e892c52-56d9-72bd-2893-b15a17cd59ae
ms.date: 06/08/2017
---


# Font2 Object (Office)

Contains font attributes (for example, font name, font size, and color) for an object.


## Example

The following example changes the formatting of the Heading 2 style in the active document to Arial and bold.


```vb
With ActiveDocument.Styles(wdStyleHeading2).Font2 
 .Name = "Arial" 
 .Italic = True 
End With 

```


## Properties



|**Name**|
|:-----|
|[Allcaps](Office.Font2.Allcaps.md)|
|[Application](Office.Font2.Application.md)|
|[AutorotateNumbers](Office.Font2.AutorotateNumbers.md)|
|[BaselineOffset](Office.Font2.BaselineOffset.md)|
|[Bold](Office.Font2.Bold.md)|
|[Caps](Office.Font2.Caps.md)|
|[Creator](Office.Font2.Creator.md)|
|[DoubleStrikeThrough](Office.Font2.DoubleStrikeThrough.md)|
|[Embeddable](Office.Font2.Embeddable.md)|
|[Embedded](Office.Font2.Embedded.md)|
|[Equalize](Office.Font2.Equalize.md)|
|[Fill](Office.Font2.Fill.md)|
|[Glow](Office.Font2.Glow.md)|
|[Highlight](Office.Font2.Highlight.md)|
|[Italic](Office.Font2.Italic.md)|
|[Kerning](Office.Font2.Kerning.md)|
|[Line](Office.Font2.Line.md)|
|[Name](Office.Font2.Name.md)|
|[NameAscii](Office.Font2.NameAscii.md)|
|[NameComplexScript](Office.Font2.NameComplexScript.md)|
|[NameFarEast](Office.Font2.NameFarEast.md)|
|[NameOther](Office.Font2.NameOther.md)|
|[Parent](Office.Font2.Parent.md)|
|[Reflection](Office.Font2.Reflection.md)|
|[Shadow](Office.Font2.Shadow.md)|
|[Size](Office.Font2.Size.md)|
|[Smallcaps](Office.Font2.Smallcaps.md)|
|[SoftEdgeFormat](Office.Font2.SoftEdgeFormat.md)|
|[Spacing](Office.Font2.Spacing.md)|
|[Strike](Office.Font2.Strike.md)|
|[StrikeThrough](Office.Font2.StrikeThrough.md)|
|[Subscript](Office.Font2.Subscript.md)|
|[Superscript](Office.Font2.Superscript.md)|
|[UnderlineColor](Office.Font2.UnderlineColor.md)|
|[UnderlineStyle](Office.Font2.UnderlineStyle.md)|
|[WordArtformat](Office.Font2.WordArtformat.md)|

## See also


#### Other resources


[Object Model Reference](http://msdn.microsoft.com/library/499c789a-aba2-0fad-649a-0ea964cd3b5e%28Office.15%29.aspx)
