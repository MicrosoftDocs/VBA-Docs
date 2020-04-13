---
title: Replacement object (Word)
keywords: vbawd10.chm2481
f1_keywords:
- vbawd10.chm2481
ms.prod: word
api_name:
- Word.Replacement
ms.assetid: 5d9615e4-f6ef-af5f-6e45-c382a88395c9
ms.date: 06/08/2017
localization_priority: Normal
---


# Replacement object (Word)

Represents the replace criteria for a find-and-replace operation. The properties and methods of the **Replacement** object correspond to the options in the **Find and Replace** dialog box.


## Remarks

Use the **Replacement** property to return a **Replacement** object. The following example replaces the next occurrence of the word "hi" with the word "hello."


```vb
With Selection.Find 
 .Text = "hi" 
 .ClearFormatting 
 .Replacement.Text = "hello" 
 .Replacement.ClearFormatting 
 .Execute Replace:=wdReplaceOne, Forward:=True 
End With
```

To find and replace formatting, set both the find text and the replace text to empty strings ("") and set the Format argument of the **Execute** method to **True**. The following example removes all the bold formatting in the active document. The **Bold** property is **True** for the **Find** object and **False** for the **Replacement** object.




```vb
With ActiveDocument.Content.Find 
 .ClearFormatting 
 .Font.Bold = True 
 .Text = "" 
 With .Replacement 
 .ClearFormatting 
 .Font.Bold = False 
 .Text = "" 
 End With 
 .Execute Format:=True, Replace:=wdReplaceAll 
End With
```


## Methods



|Name|
|:-----|
|[ClearFormatting](Word.Replacement.ClearFormatting.md)|

## Properties



|Name|
|:-----|
|[Application](Word.Replacement.Application.md)|
|[Creator](Word.Replacement.Creator.md)|
|[Font](Word.Replacement.Font.md)|
|[Frame](Word.Replacement.Frame.md)|
|[Highlight](Word.Replacement.Highlight.md)|
|[LanguageID](Word.Replacement.LanguageID.md)|
|[LanguageIDFarEast](Word.Replacement.LanguageIDFarEast.md)|
|[NoProofing](Word.Replacement.NoProofing.md)|
|[ParagraphFormat](Word.Replacement.ParagraphFormat.md)|
|[Parent](Word.Replacement.Parent.md)|
|[Style](Word.Replacement.Style.md)|
|[Text](Word.Replacement.Text.md)|

## See also


[Word Object Model Reference](overview/Word/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
