---
title: Hyperlink object (Word)
keywords: vbawd10.chm2461
f1_keywords:
- vbawd10.chm2461
ms.prod: word
api_name:
- Word.Hyperlink
ms.assetid: af785a9e-081a-e359-705f-04f490304e2e
ms.date: 06/08/2017
localization_priority: Normal
---


# Hyperlink object (Word)

Represents a hyperlink. The  **Hyperlink** object is a member of the **Hyperlinks** collection.


## Remarks

Use the  **Hyperlink** property to return a **Hyperlink** object associated with a shape (a shape can have only one hyperlink). The following example activates the hyperlink associated with the first shape in the active document.


```vb
ActiveDocument.Shapes(1).Hyperlink.Follow
```

Use  **Hyperlinks** (Index), where Index is the index number, to return a single **Hyperlink** object from a document, range, or selection. The following example activates the first hyperlink in the selection.




```vb
If Selection.HyperLinks.Count >= 1 Then 
 Selection.HyperLinks(1).Follow 
End If
```


## Methods



|Name|
|:-----|
|[AddToFavorites](Word.Hyperlink.AddToFavorites.md)|
|[CreateNewDocument](Word.Hyperlink.CreateNewDocument.md)|
|[Delete](Word.Hyperlink.Delete.md)|
|[Follow](Word.Hyperlink.Follow.md)|

## Properties



|Name|
|:-----|
|[Address](Word.Hyperlink.Address.md)|
|[Application](Word.Hyperlink.Application.md)|
|[Creator](Word.Hyperlink.Creator.md)|
|[EmailSubject](Word.Hyperlink.EmailSubject.md)|
|[ExtraInfoRequired](Word.Hyperlink.ExtraInfoRequired.md)|
|[Name](Word.Hyperlink.Name.md)|
|[Parent](Word.Hyperlink.Parent.md)|
|[Range](Word.Hyperlink.Range.md)|
|[ScreenTip](Word.Hyperlink.ScreenTip.md)|
|[Shape](Word.Hyperlink.Shape.md)|
|[SubAddress](Word.Hyperlink.SubAddress.md)|
|[Target](Word.Hyperlink.Target.md)|
|[TextToDisplay](Word.Hyperlink.TextToDisplay.md)|
|[Type](Word.Hyperlink.Type.md)|

## See also


[Word Object Model Reference](overview/Word/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]