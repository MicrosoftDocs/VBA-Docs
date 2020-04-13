---
title: Panes object (Word)
ms.prod: word
ms.assetid: 6ed6353c-9134-f47d-a108-13e84eced8ff
ms.date: 06/08/2017
localization_priority: Normal
---


# Panes object (Word)

A collection of  **Pane** objects that represent the window panes for a single window.


## Remarks

Use the **Panes** property to return the **Panes** collection. The following example splits the active window and hides the ruler for each pane.


```vb
ActiveDocument.ActiveWindow.Split = True 
For Each aPane In ActiveDocument.ActiveWindow.Panes 
 aPane.DisplayRulers = False 
Next aPane
```

Use the **Add** method or the **Split** property to add a window pane. The following example splits the active window at 20 percent of the current window size.




```vb
ActiveDocument.ActiveWindow.Panes.Add SplitVertical:=20
```

The following example splits the active window in half.




```vb
ActiveDocument.ActiveWindow.Split = True
```

You can use the **SplitSpecial** property to show comments, footnotes, or endnotes in a separate pane.

A window has more than one pane if it is split, or if the active view isn't print layout view and information such as footnotes or comments is displayed. The following example displays the footnote pane in normal view and then prompts the user to close the pane.




```vb
ActiveDocument.ActiveWindow.View.Type = wdNormalView 
If ActiveDocument.Footnotes.Count >= 1 Then 
 ActiveDocument.ActiveWindow.View.SplitSpecial = wdPaneFootnotes 
 response = _ 
 MsgBox("Do you want to close the footnotes pane?", vbYesNo) 
 If response = vbYes Then _ 
 ActiveDocument.ActiveWindow.ActivePane.Close 
End If
```


## See also



[Word Object Model Reference](overview/Word/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]