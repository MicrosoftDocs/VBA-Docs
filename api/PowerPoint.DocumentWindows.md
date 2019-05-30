---
title: DocumentWindows object (PowerPoint)
keywords: vbapp10.chm509000
f1_keywords:
- vbapp10.chm509000
ms.prod: powerpoint
api_name:
- PowerPoint.DocumentWindows
ms.assetid: 84ed4b8c-593a-8100-d4b8-158115c4e84d
ms.date: 06/08/2017
localization_priority: Normal
---


# DocumentWindows object (PowerPoint)

A collection of all the  **[DocumentWindow](PowerPoint.DocumentWindow.md)** objects that are currently open in Microsoft PowerPoint. This collection doesn't include open slide show windows, which are included in the **[SlideShowWindows](PowerPoint.SlideShowWindows.md)** collection.


## Remarks

If your Visual Studio solution includes the  **Microsoft.Office.Interop.PowerPoint** reference, this collection maps to the following types:


-  **Microsoft.Office.Interop.PowerPoint.DocumentWindows**
    

## Example

Use the [Windows](PowerPoint.Application.Windows.md) property to return the **DocumentWindows** collection. The following example tiles the open document windows.


```vb
Windows.Arrange ppArrangeTiled
```

Use the  **[NewWindow](PowerPoint.DocumentWindow.NewWindow.md)** method to create a document window and add it to the **DocumentWindows** collection. The following example creates a new window for the active presentation.




```vb
ActivePresentation.NewWindow
```

Use  **Windows** (_index_), where _index_ is the window index number, to return a single **DocumentWindow** object. The following example closes document window two.




```vb
Windows(2).Close
```


## See also


[PowerPoint Object Model Reference](overview/PowerPoint/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]