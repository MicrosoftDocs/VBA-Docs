---
title: ReflectionFormat Object (Office)
ms.prod: office
api_name:
- Office.ReflectionFormat
ms.assetid: 9684dbb3-5b99-113b-9808-1173fdd719a9
ms.date: 06/08/2017
---


# ReflectionFormat Object (Office)

Represents the reflection effect in Office graphics.


## Example

This example sets the reflection formatting for the text for the second shape on the second slide in a PowerPoint presentation:


```vb
With ActivePresentation.Slides(1).Shapes(2) 
 With .TextFrame2.TextRange.Font 
 .Size = 32 
 .Name = "Palatino" 
 .Reflection.Type = msoReflectionType6 
 End With 
End With 

```


## Properties



|**Name**|
|:-----|
|[Application](Office.ReflectionFormat.Application.md)|
|[Blur](Office.ReflectionFormat.Blur.md)|
|[Creator](Office.ReflectionFormat.Creator.md)|
|[Offset](Office.ReflectionFormat.Offset.md)|
|[Size](Office.ReflectionFormat.Size.md)|
|[Transparency](Office.ReflectionFormat.Transparency.md)|
|[Type](Office.ReflectionFormat.Type.md)|

## See also


#### Other resources


[Object Model Reference](http://msdn.microsoft.com/library/499c789a-aba2-0fad-649a-0ea964cd3b5e%28Office.15%29.aspx)
