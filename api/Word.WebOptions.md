---
title: WebOptions object (Word)
keywords: vbawd10.chm2532
f1_keywords:
- vbawd10.chm2532
ms.prod: word
api_name:
- Word.WebOptions
ms.assetid: 658ae89d-3f92-067b-1309-7fc90b257111
ms.date: 06/08/2017
localization_priority: Normal
---


# WebOptions object (Word)

Contains document-level attributes used by Microsoft Word when you save a document as a webpage or open a webpage.


## Remarks

 You can return or set attributes either at the application (global) level or at the document level. (Note that attribute values can be different from one document to another, depending on the attribute value at the time the document was saved.) Document-level attribute settings override application-level attribute settings. Application-level attributes are contained in the **DefaultWebOptions** object.

Use the **WebOptions** property to return the **WebOptions** object. The following example checks to see whether PNG (Portable Network Graphics) is allowed as an image format and then sets the _strImageFileType_ variable accordingly.




```vb
Set objAppWebOptions = ActiveDocument.WebOptions 
With objAppWebOptions 
 If .AllowPNG = True Then 
 strImageFileType = "PNG" 
 Else 
 strImageFileType = "JPG" 
 End If 
End With
```


## See also



[Word Object Model Reference](overview/Word/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]