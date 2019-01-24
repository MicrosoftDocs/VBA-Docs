---
title: ReflectionFormat object (Office)
ms.prod: office
api_name:
- Office.ReflectionFormat
ms.assetid: 9684dbb3-5b99-113b-9808-1173fdd719a9
ms.date: 01/23/2019
localization_priority: Normal
---


# ReflectionFormat object (Office)

Represents the reflection effect in Office graphics.


## Example

This example sets the reflection formatting for the text for the second shape on the second slide in a PowerPoint presentation.


```vb
With ActivePresentation.Slides(1).Shapes(2) 
 With .TextFrame2.TextRange.Font 
 .Size = 32 
 .Name = "Palatino" 
 .Reflection.Type = msoReflectionType6 
 End With 
End With 

```


## See also

- [ReflectionFormat object members](overview/Library-Reference/reflectionformat-members-office.md)
- [Object Model Reference](overview/Library-Reference/reference-object-library-reference-for-office.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]