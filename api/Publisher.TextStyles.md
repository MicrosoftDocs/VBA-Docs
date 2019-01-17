---
title: TextStyles Object (Publisher)
keywords: vbapb10.chm5963775
f1_keywords:
- vbapb10.chm5963775
ms.prod: publisher
api_name:
- Publisher.TextStyles
ms.assetid: 8a250160-0400-62e7-8301-5a5743fb2485
ms.date: 06/08/2017
localization_priority: Normal
---


# TextStyles Object (Publisher)

A collection of  **[TextStyle](Publisher.TextStyle.md)** objects that represent both the built-in and user-defined styles in a document.
 


## Example

Use the  **TextStyles** property to return the **TextStyles** collection. The following example creates a table and lists all the styles in the active publication.
 

 

```vb
Sub ListTextStyles() 
 Dim sty As TextStyle 
 Dim tbl As Table 
 Dim intRow As Integer 
 
 With ActiveDocument 
 Set tbl = .Pages(1).Shapes.AddTable(NumRows:=.TextStyles.Count, _ 
 NumColumns:=2, Left:=72, Top:=72, Width:=488, Height:=12).Table 
 For Each sty In .TextStyles 
 intRow = intRow + 1 
 With tbl.Rows(intRow) 
 .Cells(1).text = sty.Name 
 .Cells(2).text = sty.BaseStyle 
 End With 
 Next sty 
 End With 
End Sub
```

Use the  **[Add](Publisher.TextStyles.Add.md)** method to create a new user-defined style and add it to the **TextStyles** collection. The following example creates a new style and applies it to the paragraph at the cursor position.
 

 



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



|Name|
|:-----|
|[Add](Publisher.TextStyles.Add.md)|
|[Item](Publisher.TextStyles.Item.md)|

## Properties



|Name|
|:-----|
|[Application](Publisher.TextStyles.Application.md)|
|[Count](Publisher.TextStyles.Count.md)|
|[Parent](Publisher.TextStyles.Parent.md)|

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]