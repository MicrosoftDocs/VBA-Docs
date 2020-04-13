---
title: Dictionaries object (Word)
ms.prod: word
ms.assetid: 41f31292-4b3e-0d7b-c857-f6b9a0662e9a
ms.date: 06/08/2017
localization_priority: Normal
---


# Dictionaries object (Word)

A collection of  **Dictionary** objects that includes the active custom spelling dictionaries.


## Remarks

Use the **[CustomDictionaries](Word.Application.CustomDictionaries.md)** property to return the collection of currently active custom dictionaries. The following example displays the names of all the active custom dictionaries.


```vb
For Each d In CustomDictionaries 
 Msgbox d.Name 
Next d
```

Use the **[Add](Word.Dictionaries.Add.md)** method to add a new custom dictionary to the collection of active custom dictionaries. If there isn't a file with the name specified by FileName, Word creates it. The following example adds "MyCustom.dic" to the collection of custom dictionaries.




```vb
CustomDictionaries.Add FileName:="MyCustom.dic"
```

Use the **[ClearAll](Word.Dictionaries.ClearAll.md)** method to unload all custom dictionaries. Note, however, that this method doesn't delete the dictionary files. After you use this method, the number of custom dictionaries in the collection is 0 (zero). The following example clears the custom dictionaries and creates a new custom dictionary file. The new dictionary is set as the active custom dictionary, to which Word will automatically add any new words it encounters.




```vb
With CustomDictionaries 
 .ClearAll 
 .Add FileName:= "MyCustom.dic" 
 .ActiveCustomDictionary = CustomDictionaries(1) 
End With
```

Remarks

You set the custom dictionary to which new words are added by using the **[ActiveCustomDictionary](Word.Dictionaries.ActiveCustomDictionary.md)** property. If you try to set this property to a dictionary that isn't a custom dictionary, an error occurs.

The **[Maximum](Word.Dictionaries.Maximum.md)** property returns the maximum number of simultaneous custom spelling dictionaries that the application can support. For Word, this maximum is 10.


## See also



[Word Object Model Reference](overview/Word/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]