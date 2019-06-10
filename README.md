# Protected Spellcheck
Macro to easily spell check a field in a "protected" Word 2003 document.

![empty form](https://github.com/jonathancraddock/protected-spellcheck/blob/master/prot-spell-1.png "Empty protected form.")

```VBScript
Option Explicit

Sub jcCheckFieldSpelling()
    
    'v1.0 jcCheckFieldSpelling, 10/6/2019
    '...
    
    Dim userResponse As VbMsgBoxResult
    userResponse = MsgBox("Would you like to spell-check this box?", vbYesNo + vbQuestion, "Form Spellcheck")
     
    'User Clicks YES
    If userResponse = vbYes Then
    
    'Select the current form field
    ActiveDocument.Bookmarks("\Para").Select
    
    'Run built in spell check with UK English
    With Selection
        #If VBA6 Then
        .NoProofing = False
        #End If
        .LanguageID = wdEnglishUK
        .Range.CheckSpelling
    End With

    MsgBox "Spelling and Grammar check is complete.", vbInformation, "Form Spellcheck"

    'User Clicks NO
    Else
    End If

End Sub
```

![field dialog](https://github.com/jonathancraddock/protected-spellcheck/blob/master/prot-spell-2.png "Form field dialog.")

Apply to individual form fields using the "on-exit" event.
