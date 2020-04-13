---
title: ColorScheme object (PowerPoint)
keywords: vbapp10.chm537000
f1_keywords:
- vbapp10.chm537000
ms.prod: powerpoint
api_name:
- PowerPoint.ColorScheme
ms.assetid: c1945542-b628-e2b1-5114-e064f0563a01
ms.date: 06/08/2017
localization_priority: Normal
---


# ColorScheme object (PowerPoint)

Represents a color scheme, which is a set of eight colors used for the different elements of a slide, notes page, or handout, such as the title or background. (Note that the color schemes for slides, notes pages, and handouts in a presentation can be set independently.)


## Remarks

 Each color is represented by an **[RGBColor](PowerPoint.RGBColor.md)** object. The **ColorScheme** object is a member of the **[ColorSchemes](PowerPoint.ColorSchemes.md)** collection. The **ColorSchemes** collection contains all the color schemes in a presentation.

The following examples describe how to do the following:


- Return a **ColorScheme** object from the collection of all the color schemes in the presentation
    
- Return the  **ColorScheme** object attached to a specific slide or master
    
- Return the color of a single slide element from a **ColorScheme** object
    

## Example

Use  **ColorSchemes** (_index_), where _index_ is the color scheme index number, to return a single **ColorScheme** object. The following example deletes color scheme two from the active presentation.


```vb
ActivePresentation.ColorSchemes(2).Delete
```

Set the [ColorScheme](PowerPoint.Slide.ColorScheme.md)property of a **Slide**, **SlideRange**, or **Master** object to return the color scheme for one slide, a set of slides, or a master, respectively. The following example creates a color scheme based on the current slide, adds the new color scheme to the collection of standard color schemes for the presentation, and sets the color scheme for the slide master to the new color scheme. All new slides based on the master will have this color scheme.




```vb
Set newScheme = ActiveWindow.View.Slide.ColorScheme

newScheme.Colors(ppTitle).RGB = RGB(0, 150, 250)

Set newStandardScheme = _
    ActivePresentation.ColorSchemes.Add(newScheme)

ActivePresentation.SlideMaster.ColorScheme = newStandardScheme
```

Use the [Colors](PowerPoint.ColorScheme.Colors.md)method to return an **RGBColor** object that represents the color of a single slide-element type. You can set an **RGBColor** object to another **RGBColor** object, or you can use the [RGB](PowerPoint.ColorFormat.RGB.md)property to set or return the explicit red-green-blue (RGB) value for an **RGBColor** object. The following example sets the background color in color scheme one to red and sets the title color to the title color that's defined for color scheme two.




```vb
With ActivePresentation.ColorSchemes

    .Item(1).Colors(ppBackground).RGB = RGB(255, 0, 0)

    .Item(1).Colors(ppTitle) = .Item(2).Colors(ppTitle)

End With
```


## See also


[PowerPoint Object Model Reference](overview/PowerPoint/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]