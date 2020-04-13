---
title: Template object (Word)
keywords: vbawd10.chm2410
f1_keywords:
- vbawd10.chm2410
ms.prod: word
api_name:
- Word.Template
ms.assetid: 47d1d92d-bba9-3f2a-9c71-22ac43159bd3
ms.date: 06/08/2017
localization_priority: Normal
---


# Template object (Word)

Represents a document template. The **Template** object is a member of the **[Templates](Word.templates.md)** collection. The **Templates** collection includes all the available **Template** objects.


## Remarks

Use  **Templates** (Index), where Index is the template name or the index number, to return a single **Template** object. The following example saves the Memo2.dot template if it is in the **Templates** collection.


```vb
For Each aTemp In Templates 
 If LCase(aTemp.Name) = "memo2.dot" Then aTemp.Save 
Next aTemp
```

The index number represents the position of the template in the **Templates** collection. The following example opens the first template in the **Templates** collection.




```vb
Templates(1).OpenAsDocument
```

The **Add** method is not available for the **Templates** collection. Instead, you can add a template to the **Templates** collection by doing any of the following:


- Using the **Open** method with the **Documents** collection to open a document based on a template or a template
    
- Using the **Add** method with the **Documents** collection to open a new document based on a template
    
- Using the **Add** method with the **Addins** collection to load a global template
    
- Using the **AttachedTemplate** property with the **Document** object to attach a template to a document
    
Use the **NormalTemplate** property to return a template object that refers to the Normal template. Use the **AttachedTemplate** property to return the template attached to the specified document.

Use the **DefaultFilePath** property to return or set the location of user or workgroup templates (that is, the folder where you want to store these templates). The following example displays the user template folder from the **File Locations** tab in the **Options** dialog box (**Tools** menu).




```vb
MsgBox Options.DefaultFilePath(wdUserTemplatesPath)
```


## Methods



|Name|
|:-----|
|[OpenAsDocument](Word.Template.OpenAsDocument.md)|
|[Save](Word.Template.Save.md)|

## Properties



|Name|
|:-----|
|[Application](Word.Template.Application.md)|
|[BuildingBlockEntries](Word.Template.BuildingBlockEntries.md)|
|[BuildingBlockTypes](Word.Template.BuildingBlockTypes.md)|
|[BuiltInDocumentProperties](Word.Template.BuiltInDocumentProperties.md)|
|[Creator](Word.Template.Creator.md)|
|[CustomDocumentProperties](Word.Template.CustomDocumentProperties.md)|
|[FarEastLineBreakLanguage](Word.Template.FarEastLineBreakLanguage.md)|
|[FarEastLineBreakLevel](Word.Template.FarEastLineBreakLevel.md)|
|[FullName](Word.Template.FullName.md)|
|[JustificationMode](Word.Template.JustificationMode.md)|
|[KerningByAlgorithm](Word.Template.KerningByAlgorithm.md)|
|[LanguageID](Word.Template.LanguageID.md)|
|[LanguageIDFarEast](Word.Template.LanguageIDFarEast.md)|
|[ListTemplates](Word.Template.ListTemplates.md)|
|[Name](Word.Template.Name.md)|
|[NoLineBreakAfter](Word.Template.NoLineBreakAfter.md)|
|[NoLineBreakBefore](Word.Template.NoLineBreakBefore.md)|
|[NoProofing](Word.Template.NoProofing.md)|
|[Parent](Word.Template.Parent.md)|
|[Path](Word.Template.Path.md)|
|[Saved](Word.Template.Saved.md)|
|[Type](Word.Template.Type.md)|
|[VBProject](Word.Template.VBProject.md)|

## See also


[Word Object Model Reference](overview/Word/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]