---
title: Slides.Add method (PowerPoint)
api_name:
- PowerPoint.Slides.AddSlide
---


# Slides.Add method (PowerPoint)

Creates a new slide, adds it to the **[Slides](PowerPoint.Slides.md)** collection, and returns the slide.

## Syntax

_expression_. `AddSlide`( `_Index_`, `_Layout_` )

 _expression_ An expression that returns a [Slides](PowerPoint.Slides.md) object.

## Parameters


|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Index_|Required|**Int**|The index of the slide to be added.|
| _Layout_|Required|**PpSlideLayout**|The layout of the slide.|

## Return value

Slide

## Example

The following example shows how to use the **Add** method to add a new slide to the **Slides** collection. It adds a new slide in index position 1 with the blank layout.

```vb
Public Sub Add_Example() 
 
    Dim pptSlide As Slide 
    Set pptSlide = ActivePresentation.Slides.AddSlide(1, ppLayoutBlank) 
 
End Sub
```

## Remarks

See [PpSlideLayout enumeration](PowerPoint.PpSlideLayout.md) for all available layouts. For custom slide layouts, use [Slides.AddSlide](PowerPoint.Slides.AddSlide.md) instead.

This method is hidden from the VBA Object Browser, but is still available to use.

If your Visual Studio solution includes the **Microsoft.Office.Interop.PowerPoint** reference, this method maps to the following type:

- **Microsoft.Office.Interop.PowerPoint.Slides.Add(int, Microsoft.Office.Interop.PowerPoint.PpSlideLayout)**

## See also

[Slides Object](PowerPoint.Slides.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
