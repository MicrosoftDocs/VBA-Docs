---
title: FileConverters object (Word)
ms.prod: word
ms.assetid: b9b8fc53-1c8e-224d-726a-4edf172ca647
ms.date: 06/08/2017
localization_priority: Normal
---


# FileConverters object (Word)

A collection of  **FileConverter** objects that represent all the file converters available for opening and saving files.


## Remarks

Use the  **FileConverters** property to return the **FileConverters** collection. The following example determines whether a WordPerfect 6.0 converter is available.


```vb
For Each conv In FileConverters 
 If conv.FormatName = "WordPerfect 6.x" Then 
 MsgBox "WordPerfect 6.0 converter is installed" 
 End if 
Next conv
```

The  **Add** method isn't available for the **FileConverters** collection. **[FileConverter](Word.FileConverter.md)** objects are added during installation of Microsoft Office or by installing supplemental converters.

Use  **FileConverters** (Index), where Index is a class name or index number, to return a single **[FileConverter](Word.FileConverter.md)** object. The following example displays the extensions associated with the Microsoft Excel worksheet converter.




```vb
MsgBox FileConverters("MSBiff").Extensions
```

The index number represents the position of the file converter in the  **FileConverters** collection. The following example displays the format name of the first file converter.




```vb
MsgBox FileConverters(1).FormatName
```

File converters for saving documents are listed in the  **Save As** dialog box. File converters for opening documents appear in a dialog box if the **Confirm conversion at Open** check box is selected on the **General** tab in the **Options** dialog box (**Tools** menu).


## See also


[Word Object Model Reference](overview/Word/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]