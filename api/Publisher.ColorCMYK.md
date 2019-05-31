---
title: ColorCMYK object (Publisher)
keywords: vbapb10.chm2686975
f1_keywords:
- vbapb10.chm2686975
ms.prod: publisher
api_name:
- Publisher.ColorCMYK
ms.assetid: e1a39f6f-f440-e375-4f8c-e81093e5a451
ms.date: 05/31/2019
localization_priority: Normal
---


# ColorCMYK object (Publisher)

Represents a cyan-magenta-yellow-black (CMYK) color value.
 
## Remarks

Use the **[CMYK](publisher.colorformat.cmyk.md)** property of the **ColorFormat** object to return a **ColorCMYK** object.

Use the **Cyan**, **Magenta**, **Yellow**, and **Black** properties to individually set each of the four colors in the CMYK color value. Use the **SetCMYK** method to set all four colors at once.

## Example

The following example retrieves the CMYK color value of shape one's fill and changes it to another CMYK color value.

```vb
Dim cmykColor As ColorCMYK Set cmykColor = ActiveDocument.Pages(1).Shapes(1).Fill.ForeColor.CMYK cmykColor.SetCMYK Cyan:=0, Magenta:=255, Yellow:=255, Black:=50
```


## Methods

- [SetCMYK](Publisher.ColorCMYK.SetCMYK.md)

## Properties

- [Application](Publisher.ColorCMYK.Application.md)
- [Black](Publisher.ColorCMYK.Black.md)
- [Cyan](Publisher.ColorCMYK.Cyan.md)
- [Magenta](Publisher.ColorCMYK.Magenta.md)
- [Parent](Publisher.ColorCMYK.Parent.md)
- [Yellow](Publisher.ColorCMYK.Yellow.md)

## See also

- [Publisher Object Model Reference](overview/publisher/object-model.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]