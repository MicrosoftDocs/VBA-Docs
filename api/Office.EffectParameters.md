---
title: EffectParameters object (Office)
ms.prod: office
api_name:
- Office.EffectParameters
ms.assetid: 9b0dfcf1-96fa-bc9a-6fef-38518ab1c558
ms.date: 01/08/2019
localization_priority: Normal
---


# EffectParameters object (Office)

Represents a collection of **[EffectParameter](office.effectparameter.md)** objects.


## Remarks

Picture Effects are processed as a chain composed of individual items that are applied in sequence to create the final composited image. An Effects chain will allow an effect to be added to the chain, reordered, or removed from the chain. Effect parameters specify properties of those effects.


## Example

The following code sets several Picture Effect fill properties on a shape in a Microsoft PowerPoint slide.


```vb
Sub PictureEffectSample() 
' Setup a slide with one picture shape. 
With ActivePresentation.Slides(1).Shapes(1).Fill.PictureEffects 
 
 ' Insert a 150% Saturation effect. 
 .Insert(msoEffectSaturation).EffectParameters(1).Value = 1.5 
 
 ' Insert Brightness/Contrast effect and set values to -50% Brightness and +25% Contrast. 
 Dim brightnessContrast As PictureEffect 
 Set brightnessContrast = .Insert(msoEffectBrightnessContrast) 
 brightnessContrast.EffectParameters(1).Value = -0.5 
 brightnessContrast.EffectParameters(2).Value = 0.25 
 
 ' Remove all Picture effects. 
 While .Count > 0 
 .Delete (1) 
 Wend 
 
End With 
End Sub
```




## See also

- [EffectParameters object members](overview/library-reference/effectparameters-members-office.md)
- [Object Model Reference](overview/library-reference/reference-object-library-reference-for-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]