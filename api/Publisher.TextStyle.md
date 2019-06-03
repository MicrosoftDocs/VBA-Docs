---
title: TextStyle object (Publisher)
keywords: vbapb10.chm6029311
f1_keywords:
- vbapb10.chm6029311
ms.prod: publisher
api_name:
- Publisher.TextStyle
ms.assetid: 163ab726-ac44-07d1-ab7b-50061037cc77
ms.date: 06/04/2019
localization_priority: Normal
---


# TextStyle object (Publisher)

Represents a single built-in or user-defined style. The **TextStyle** object includes style attributes (font, font style, paragraph spacing, and so on) as properties of the **TextStyle** object. 

The **TextStyle** object is a member of the **[TextStyles](Publisher.TextStyles.md)** collection. The **TextStyles** collection includes all the styles in the specified document.
 

## Remarks

Use **TextStyles** (_index_), where _index_ is the text style number or name, to return a single **TextStyle** object. You must exactly match the spelling and spacing of the style name, but not necessarily its capitalization.

Use the **[TextStyles.Add](Publisher.TextStyles.Add.md)** method to create a new style. 

To apply a style to a range, paragraph, or multiple paragraphs, set the **[ParagraphFormat.TextStyle](Publisher.ParagraphFormat.TextStyle.md)** property to a user-defined or built-in style name. 


## Example
 
The following example displays the style name and base style of the first style in the **TextStyles** collection.

```vb
Sub BaseStyleName() 
 With ActiveDocument.TextStyles(1) 
 MsgBox "Style name= " & .Name _ 
 & vbCr & "Base style= " & .BaseStyle 
 End With 
End Sub
```

<br/>

The following example creates a new style and applies it to the paragraph at the cursor position.

```vb
Sub ApplyTextStyle() 
 Dim styNew As TextStyle 
 Dim fntStyle As Font 
 
 'Create a new style 
 Set styNew = ActiveDocument.TextStyles.Add(StyleName:="NewStyle") 
 Set fntStyle = styNew.Font 
 
 'Format the Font object 
 With fntStyle 
 .Name = "Tahoma" 
 .Size = 20 
 .Bold = msoTrue 
 End With 
 
 'Apply the Font object formatting to the new style 
 styNew.Font = fntStyle 
 
 'Apply the new style to the selected paragraph 
 Selection.TextRange.ParagraphFormat.TextStyle = "NewStyle" 
End Sub
```


## Methods

- [Delete](Publisher.TextStyle.Delete.md)

## Properties

- [Application](Publisher.TextStyle.Application.md)
- [BaseStyle](Publisher.TextStyle.BaseStyle.md)
- [Description](Publisher.TextStyle.Description.md)
- [Font](Publisher.TextStyle.Font.md)
- [Name](Publisher.TextStyle.Name.md)
- [NextParagraphStyle](Publisher.TextStyle.NextParagraphStyle.md)
- [ParagraphFormat](Publisher.TextStyle.ParagraphFormat.md)
- [Parent](Publisher.TextStyle.Parent.md)

## See also

- [Publisher Object Model Reference](overview/publisher/object-model.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]