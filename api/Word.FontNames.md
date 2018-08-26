---
title: FontNames Object (Word)
ms.prod: word
api_name:
- Word.FontNames
ms.assetid: d3a9a52f-b441-ac63-3e12-25dbf1022f38
ms.date: 06/08/2017
---


# FontNames Object (Word)

Represents a list of the names of all the available fonts.


## Remarks

Use the  **FontNames** , **LandscapeFontNames** , or **PortraitFontNames** property to return the **FontNames** object. The following example displays the number of portrait fonts available.


```vb
<<<<<<< HEAD
MsgBox PortraitFontNames.Count &; " fonts available"
=======
MsgBox PortraitFontNames.Count & " fonts available"
>>>>>>> master
```

This example lists all the font names in the  **FontNames** object at the end of the active document.




```vb
For Each aFont In FontNames 
<<<<<<< HEAD
 ActiveDocument.Range.InsertAfter aFont &; vbCr 
=======
 ActiveDocument.Range.InsertAfter aFont & vbCr 
>>>>>>> master
Next aFont
```

Use  **FontNames** (Index), where Index is the index number, to return the name of a font. The following example displays the first font name in the **FontNames** object.




```vb
MsgBox FontNames(1)
```


 **Note**  You cannot add names to or remove names from the list of available font names.


## See also



[Word Object Model Reference](./overview/Word/object-model.md)

