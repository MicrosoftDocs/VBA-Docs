---
title: Zooms object (Word)
keywords: vbawd10.chm2471
f1_keywords:
- vbawd10.chm2471
ms.prod: word
ms.assetid: 1a4d5375-ad91-1eb9-77cb-4a6f8dcc3eb8
ms.date: 06/08/2017
localization_priority: Normal
---


# Zooms object (Word)

A collection of  **[Zoom](Word.Zoom.md)** objects that represents the magnification options for each view (such as outline, normal, or print layout).


## Remarks

Use the **Zooms** property to return the **Zooms** collection. The following example sets the zoom percentage for the active window to 100 percent in Normal view.


```vb
ActiveDocument.ActiveWindow.ActivePane _ 
 .Zooms(wdNormalView).Percentage = 100
```

The **Add** method isn't available for the **Zooms** collection. The **Zooms** collection includes a single **Zoom** object for each of the various view types (such as outline, normal, or page layout). You cannot enumerate the **Zooms** collection by using a **For Each** loop.

Use  **Zooms** (index), where index identifies the view type, to return a single **Zoom** object. The view type specified by index can be one of the following **[WdViewType](Word.WdViewType.md)** constants: **wdMasterView**, **wdNormalView**, **wdOutlineView**, **wdPrintPreview**, **wdPrintView**, or **wdWebView**. The following example sets the magnification for the active window so that an entire page is visible.




```vb
ActiveDocument.ActiveWindow.ActivePane _ 
 .Zooms(wdPrintView).PageFit = wdPageFitFullPage
```

You can also use the **Zoom** property of the **View** object to return a single **Zoom** object. The following example sets the zoom percentage for the active window to 110 percent.




```vb
ActiveDocument.ActiveWindow.View.Zoom.Percentage = 110
```


## See also


[Word Object Model Reference](overview/Word/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]