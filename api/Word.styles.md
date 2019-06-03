---
title: Styles object (Word)
keywords: vbawd10.chm2349
f1_keywords:
- vbawd10.chm2349
ms.prod: word
ms.assetid: bc4688ce-5055-c135-a656-e58e31d8be42
ms.date: 06/08/2017
localization_priority: Normal
---


# Styles object (Word)

A collection of  **Style** objects that represent both the built-in and user-defined styles in a document.


## Remarks

Use the  **Styles** property to return the **Styles** collection. The following example deletes all user-defined styles in the active document.


```vb
For Each sty In ActiveDocument.Styles 
 If sty.BuiltIn = False Then sty.Delete 
Next sty
```

Use the  **Add** method to create a new user-defined style and add it to the **Styles** collection. The following example adds a new character style named "Introduction" and makes it 12-point Arial, with bold and italic formatting. The example then applies this new character style to the selection.




```vb
Set myStyle = ActiveDocument.Styles.Add(Name:="Introduction", _ 
 Type:=wdStyleTypeCharacter) 
With myStyle.Font 
 .Bold = True 
 .Italic = True 
 .Name = "Arial" 
 .Size = 12 
End With 
Selection.Range.Style = "Introduction"
```

Use  **Styles** (Index), where Index is the style name, a **WdBuiltinStyle** constant or index number, to return a single **[Style](Word.Style.md)** object. You must exactly match the spelling and spacing of the style name, but not necessarily its capitalization. The following example modifies the font of the user-defined style named "Color" in the active document.




```vb
ActiveDocument.Styles("Color").Font.Name = "Arial"
```

The following example sets the built-in Heading 1 style to not be bold.




```vb
ActiveDocument.Styles(wdStyleHeading1).Font.Bold = False
```

The style index number represents the position of the style in the alphabetically sorted list of style names. Note that  `Styles(1)` is the first style in the alphabetical list. The following example displays the base style and style name of the first style in the **Styles** collection.




```vb
MsgBox "Base style= " _ 
 & ActiveDocument.Styles(1).BaseStyle & vbCr _ 
 & "Style name= " & ActiveDocument.Styles(1).NameLocal
```

The  **Styles** object is not available from the **Template** object. However, you can use the **OpenAsDocument** method to open a template as a document so that you can modify styles in the template. The following example changes the formatting of the Heading 1 style in the template attached to the active document.




```vb
Set aDoc = ActiveDocument.AttachedTemplate.OpenAsDocument 
With aDoc 
 .Styles(wdStyleHeading1).Font.Name = "Arial" 
 .Close SaveChanges:=wdSaveChanges 
End With
```

Use the  **OrganizerCopy** method to copy styles between documents and templates. Use the **UpdateStyles** method to update the styles in the active document to match the style definitions in the attached template.


## Methods



|Name|
|:-----|
|[Add](Word.Styles.Add.md)|
|[Item](Word.Styles.Item.md)|

## Properties



|Name|
|:-----|
|[Application](Word.Styles.Application.md)|
|[Count](Word.Styles.Count.md)|
|[Creator](Word.Styles.Creator.md)|
|[Parent](Word.Styles.Parent.md)|

## See also


[Word Object Model Reference](overview/Word/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]