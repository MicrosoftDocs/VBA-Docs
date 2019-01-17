---
title: Match entries in a list
keywords: fm20.chm5225258
f1_keywords:
- fm20.chm5225258
ms.prod: office
ms.assetid: 29926096-657b-ea66-e673-a0f82e6e5026
ms.date: 12/29/2018
localization_priority: Normal
---


# Match entries in a list

Microsoft Forms provides three ways to match a value entered by the user with an entry that exists in the list of a **[ListBox](../../reference/user-interface-help/listbox-control.md)** or **[ComboBox](../../reference/user-interface-help/combobox-control.md)**:

- **No matching** provides no assistance in matching a user's typed entry to an entry in the list.
    
- **First letter** compares the most recently-typed letter to the first letter of each entry in the list. The first match in the list is selected.
    
- **Complete** compares the user's entry and tries to find an exact match in an entry from the list.
    
The matching feature resets after two seconds (six seconds if you are using East Asia settings). For example, if you have a list of the 50 states and you type "CO" quickly, you will find "Colorado." But if you type "CO" slowly, you will find "Ohio" because the auto-complete search resets between letters.

If you choose **Complete** matching, it is a good idea to sort the list entries alphabetically (you can use the **[TextColumn](../../reference/user-interface-help/textcolumn-property.md)** property to do this). If the list is not sorted alphabetically, matching may not work correctly. For example, if the list includes Alabama, Louisiana, and Alaska in that order, "Alabama" will be considered a complete match if the user types "ala." In fact, this result is ambiguous because there are two entries in the list that could match what the user entered. Sorting alphabetically eliminates this ambiguity.

## See also

- [Microsoft Forms reference](../../reference/user-interface-help/reference-microsoft-forms.md)
- [Microsoft Forms conceptual topics](../../reference/user-interface-help/concepts-microsoft-forms.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]