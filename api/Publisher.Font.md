---
title: Font object (Publisher)
keywords: vbapb10.chm5439487
f1_keywords:
- vbapb10.chm5439487
ms.prod: publisher
api_name:
- Publisher.Font
ms.assetid: 992fda94-2820-d665-0d78-efd4b5434731
ms.date: 05/31/2019
localization_priority: Normal
---


# Font object (Publisher)

Contains font attributes (font name, font size, color, and so on) for an object.

## Remarks

Use the **[TextStyle.Font](Publisher.TextStyle.Font.md)** property to return the **Font** object. 

## Example

The following instruction applies bold formatting to the selection.

```vb
Sub BoldText() 
 Selection.TextRange.Font.Bold = True 
End Sub
```

<br/>

The following example formats the first paragraph in the active publication as 24-point Arial and italic.

```vb
Sub FormatText() 
 Dim txtRange As TextRange 
 Set txtRange = ActiveDocument.Pages(1).Shapes(1).TextFrame.TextRange 
 With txtRange.Font 
 .Bold = True 
 .Name = "Arial" 
 .Size = 24 
 End With 
End Sub
```

<br/>

The following example changes the formatting of the Heading 2 style in the active publication to Arial and bold.

```vb
Sub FormatStyle() 
 With ActiveDocument.TextStyles("Normal").Font 
 .Name = "Tahoma" 
 .Italic = True 
 .Size = 15 
 End With 
End Sub
```

<br/>

You can also duplicate a **Font** object by using the **[TextRange.Duplicate](Publisher.TextRange.Duplicate.md)** property. The following example creates a new character style with the character formatting from the selection in addition to italic formatting. The formatting of the selection is not changed.

```vb
Sub DuplicateFont() 
 Dim fntNew As Font 
 Set fntNew = Selection.TextRange.Font.Duplicate 
 fntNew.Italic = True 
 ActiveDocument.TextStyles.Add(StyleName:="Italics").Font = fntNew 
End Sub
```


## Methods

- [Duplicate](Publisher.Font.Duplicate.md)
- [GetScriptName](Publisher.Font.GetScriptName.md)
- [Grow](Publisher.Font.Grow.md)
- [Reset](Publisher.Font.Reset.md)
- [SetScriptName](Publisher.Font.SetScriptName.md)
- [Shrink](Publisher.Font.Shrink.md)

## Properties

- [AllCaps](Publisher.Font.AllCaps.md)
- [Application](Publisher.Font.Application.md)
- [AttachedToText](Publisher.Font.AttachedToText.md)
- [AutomaticPairKerningThreshold](Publisher.Font.AutomaticPairKerningThreshold.md)
- [Bold](Publisher.Font.Bold.md)
- [BoldBi](Publisher.Font.BoldBi.md)
- [ContextualAlternates](Publisher.Font.ContextualAlternates.md)
- [DiacriticColor](Publisher.Font.DiacriticColor.md)
- [ExpandUsingKashida](Publisher.Font.ExpandUsingKashida.md)
- [Fill](Publisher.font.fill.md)
- [Glow](Publisher.font.glow.md)
- [Italic](Publisher.Font.Italic.md)
- [ItalicBi](Publisher.Font.ItalicBi.md)
- [Kerning](Publisher.Font.Kerning.md)
- [Ligature](Publisher.font.ligature.md)
- [Line](Publisher.font.line.md)
- [Name](Publisher.Font.Name.md)
- [NumberStyle](Publisher.font.numberstyle.md)
- [Parent](Publisher.Font.Parent.md)
- [Position](Publisher.Font.Position.md)
- [Reflection](Publisher.font.reflection.md)
- [Scaling](Publisher.Font.Scaling.md)
- [Size](Publisher.Font.Size.md)
- [SizeBi](Publisher.Font.SizeBi.md)
- [SmallCaps](Publisher.Font.SmallCaps.md)
- [StrikeThrough](Publisher.font.strikethrough.md)
- [StylisticAlternates](Publisher.Font.StylisticAlternates.md)
- [StylisticSets](Publisher.Font.StylisticSets.md)
- [SubScript](Publisher.Font.SubScript.md)
- [SuperScript](Publisher.Font.SuperScript.md)
- [Swash](Publisher.Font.Swash.md)
- [TextShadow](Publisher.font.textshadow.md)
- [ThreeD](Publisher.font.threed.md)
- [Tracking](Publisher.Font.Tracking.md)
- [TrackingPreset](Publisher.Font.TrackingPreset.md)
- [Underline](Publisher.Font.Underline.md)
- [UseDiacriticColor](Publisher.Font.UseDiacriticColor.md)

## See also

- [Publisher Object Model Reference](overview/publisher/object-model.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]