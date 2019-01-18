---
title: AutoTextEntries object (Word)
ms.prod: word
ms.assetid: 4e4d92b3-d259-84b7-061f-82065e177c29
ms.date: 06/08/2017
localization_priority: Normal
---


# AutoTextEntries object (Word)

A collection of  **[AutoCorrectEntry](Word.AutoCorrectEntry.md)** objects that represent the AutoText entries in a template. The **AutoTextEntries** collection includes all the entries listed on the **AutoText** tab in the **AutoCorrect** dialog box.


## Remarks

Use the  **AutoTextEntries** property to return the **AutoTextEntries** collection. The following example determines whether an **[AutoTextEntry](Word.AutoTextEntry.md)** object named "test" is in the **AutoTextEntries** collection.


```vb
For Each i In NormalTemplate.AutoTextEntries 
 If LCase(i.Name) = "test" Then MsgBox "AutoText entry exists" 
Next i
```

Use the  **[Add](Word.AutoTextEntries.Add.md)** method to add an AutoText entry to the **AutoTextEntries** collection. The following example adds an AutoText entry named "Blue" based on the text of the selection.




```vb
NormalTemplate.AutoTextEntries.Add Name:="Blue", _ 
 Range:=Selection.Range
```

Use  **AutoTextEntries** (index), where index is the AutoText entry name or index number, to return a single **AutoTextEntry** object. You must exactly match the spelling (but not necessarily the capitalization) of the name, as it is shown on the **AutoText** tab in the **AutoCorrect** dialog box. The following example sets the value of an existing AutoText entry named "cName."




```vb
NormalTemplate.AutoTextEntries("cName").Value = _ 
 "The Johnson Company"
```

The following example displays the name and value of the first AutoText entry in the template attached to the active document.




```vb
Set myTemplate = ActiveDocument.AttachedTemplate 
MsgBox "Name = " & myTemplate.AutoTextEntries(1).Name & vbCr _ 
 & "Value " & myTemplate.AutoTextEntries(1).Value
```


## See also


[Word Object Model Reference](./overview/Word/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]