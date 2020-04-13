---
title: ListBox.MatchEntry Property (Outlook Forms Script)
ms.prod: outlook
ms.assetid: 8a4b99c5-2c22-a02a-e79b-6d1af40f7bfc
ms.date: 06/08/2017
localization_priority: Normal
---


# ListBox.MatchEntry Property (Outlook Forms Script)

Returns or sets an **Integer** that indicates how a **[ListBox](Outlook.listbox.md)** searches its list as the user types. Read/write.


## Syntax

_expression_.**MatchEntry**

_expression_ A variable that represents a **ListBox** object.


## Remarks

The settings for  **MatchEntry** are:



|Value|Description|
|:-----|:-----|
|0|Basic matching. The control searches for the next entry that starts with the character entered. Repeatedly typing the same letter cycles through all entries beginning with that letter.|
|1|Extended matching. As each character is typed, the control searches for an entry matching all characters entered (default).|
|2|No matching.|

The **MatchEntry** property searches entries from the **[TextColumn](Outlook.listbox.textcolumn.md)** property of a **ListBox** or **ComboBox**.

The control searches the column identified by  **TextColumn** for an entry that matches the user's typed entry. Upon finding a match, the row containing the match is selected, the contents of the column are displayed, and the contents of its **[BoundColumn](Outlook.listbox.boundcolumn.md)** property become the value of the control. If the match is unambiguous, finding the match initiates the **[Click](Outlook.listbox.click.md)** event.

The control initiates the  **Click** event as soon as the user types a sequence of characters that match exactly one entry in the list. As the user types, the entry is compared with the current row in the list and with the next row in the list. When the entry matches only the current row, the match is unambiguous.

In Microsoft Forms, this is true regardless of whether the list is sorted. This means the control finds the first occurrence that matches the entry, based on the order of items in the list. For example, entering either "abc" or "bc" will initiate the  **Click** event for the following list:




```vb
abcde 
bcdef 
abcxyz 
bchij
```

Note that in either case, the matched entry is not unique; however, it is sufficiently different from the adjacent entry that the control interprets the match as unambiguous and initiates the  **Click** event.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]