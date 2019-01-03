---
title: EffectParameter.Value property (Office)
ms.prod: office
api_name:
- Office.EffectParameter.Value
ms.assetid: 45bf51fe-c049-1c8e-cc3b-fdbd5d6d7157
ms.date: 06/08/2017
---


# EffectParameter.Value property (Office)

Retrieves or sets the value of the  **EffectParameter** object. Read/write


## Syntax

 _expression_. `Value`

 _expression_ An expression that returns a [EffectParameter](./Office.EffectParameter.md) object.


## Example

The following code sets the first parameter of the  **PictureEffect** object as color temperature.


```vb
Dim picEffect As PictureEffect 
 
picEffect.EffectParameters(1).Value = MsoPictureEffectType.msoEffectColorTemperature
```


## See also


[EffectParameter Object](Office.EffectParameter.md)



[EffectParameter Object Members](./overview/Library-Reference/effectparameter-members-office.md)

