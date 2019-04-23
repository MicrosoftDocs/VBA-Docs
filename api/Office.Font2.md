---
title: Font2 object (Office)
ms.prod: office
api_name:
- Office.Font2
ms.assetid: 8e892c52-56d9-72bd-2893-b15a17cd59ae
ms.date: 01/09/2019
localization_priority: Normal
---


# Font2 object (Office)

Contains font attributes (for example, font name, font size, and color) for an object.


## Example

The following example changes the formatting of the Heading 2 style in the active document to Arial and bold.


```vb
With ActiveDocument.Styles(wdStyleHeading2).Font2 
 .Name = "Arial" 
 .Italic = True 
End With 

```


## See also

- [Font2 object members](overview/library-reference/font2-members-office.md)
- [Object Model Reference](overview/Library-Reference/reference-object-library-reference-for-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]