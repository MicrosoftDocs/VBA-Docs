---
title: Create a Rule to Move Specific Emails to a Folder
ms.prod: outlook
ms.assetid: e72fa307-8224-c2d2-1318-a18cd8e9f22f
ms.date: 06/08/2017
localization_priority: Normal
---


# Create a Rule to Move Specific Emails to a Folder

This topic shows a code sample in Visual Basic for Applications (VBA) that uses the  **Rules** object model to create a rule. The code sample uses the **[RuleAction](../../../api/Outlook.RuleAction.md)** and **[RuleCondition](../../../api/Outlook.RuleCondition.md)** objects to specify a rule that moves messages from a specific sender to a specific folder, unless the message contains certain terms in the subject. Note that the code sample assumes that there already exists a folder named "Dan" under the Inbox.

The following describes the steps used to create the rule:

1. Specify the target folder  `oMoveTarget` to move specific messages as determined by the condition and exception condition. The target folder is a subfolder named "Dan" under the Inbox, and is assumed to already exist.
    
2. Use  **[Store.GetRules](../../../api/Outlook.Store.GetRules.md)** to obtain a set of all the rules in the current session.
    
3. Using the  **[Rules](../../../api/Outlook.Rules.md)** collection returned from the last step, use **[Rules.Create](../../../api/Outlook.Rules.Create.md)** to add a new rule. The new rule specifies some action upon receiving a message, so it is of type **olRuleReceive**.
    
4. Using the  **[Rule](../../../api/Outlook.Rule.md)** object returned from the last step, use the **[RuleConditions.From](../../../api/Outlook.RuleConditions.From.md)** property to obtain a **[ToOrFromRuleCondition](../../../api/Outlook.ToOrFromRuleCondition.md)** object, `oFromCondition`.  `oFromCondition` specifies the condition for the rule: when a message is from `Dan Wilson`. 
    
5. Using the same  **Rule** object, use the **[RuleActions.MoveToFolder](../../../api/Outlook.RuleActions.MoveToFolder.md)** property to obtain a **[MoveOrCopyRuleAction](../../../api/Outlook.MoveOrCopyRuleAction.md)** object, `oMoveRuleAction`.  `oMoveRuleAction` specifies the action for the rule: move the message to the target folder "Dan".
    
6. Using the same  **Rule** object, use the **[RuleConditions.Subject](../../../api/Outlook.RuleConditions.Subject.md)** property to obtain a **[TextRuleCondition](../../../api/Outlook.TextRuleCondition.md)** object, `oExceptSubject`.  `oExceptSubject` specifies the exception condition: if the subject contains the terms "fun" or "chat", then do not apply the rule to move the message to the folder "Dan".
    
7. Use  **[Rules.Save](../../../api/Outlook.Rules.Save.md)** to save the new rule together with the rest of the rules for the current store.
    



```vb
Sub CreateRule() 
    Dim colRules As Outlook.Rules 
    Dim oRule As Outlook.Rule 
    Dim colRuleActions As Outlook.RuleActions 
    Dim oMoveRuleAction As Outlook.MoveOrCopyRuleAction 
    Dim oFromCondition As Outlook.ToOrFromRuleCondition 
    Dim oExceptSubject As Outlook.TextRuleCondition 
    Dim oInbox As Outlook.Folder 
    Dim oMoveTarget As Outlook.Folder 
 
    'Specify target folder for rule move action 
    Set oInbox = Application.Session.GetDefaultFolder(olFolderInbox) 
    'Assume that target folder already exists 
    Set oMoveTarget = oInbox.Folders("Dan") 
     
    'Get Rules from Session.DefaultStore object 
    Set colRules = Application.Session.DefaultStore.GetRules() 
     
    'Create the rule by adding a Receive Rule to Rules collection 
    Set oRule = colRules.Create("Dan's rule", olRuleReceive) 
 
    'Specify the condition in a ToOrFromRuleCondition object 
    'Condition is if the message is from "Dan Wilson" 
    Set oFromCondition = oRule.Conditions.From 
    With oFromCondition 
        .Enabled = True 
        .Recipients.Add ("Dan Wilson") 
        .Recipients.ResolveAll 
    End With 
 
    'Specify the action in a MoveOrCopyRuleAction object 
    'Action is to move the message to the target folder 
    Set oMoveRuleAction = oRule.Actions.MoveToFolder 
    With oMoveRuleAction 
        .Enabled = True 
        .Folder = oMoveTarget 
    End With 
 
    'Specify the exception condition for the subject in a TextRuleCondition object 
    'Exception condition is if the subject contains "fun" or "chat" 
    Set oExceptSubject = _ 
        oRule.Exceptions.Subject 
    With oExceptSubject 
        .Enabled = True 
        .Text = Array("fun", "chat") 
    End With 
 
    'Update the server and display progress dialog 
    colRules.Save 
End Sub 
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
