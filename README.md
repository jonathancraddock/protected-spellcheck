# Protected Spellcheck

### Scenario

External third-parties provide users with a Word document/template (in Word 2003 format) which they are required to complete and submit. The document is protected and users enter text into plain text form fields. The process can be frustrating since these legacy fields offer restricted editing and it is not possible to spell-check their content.

Workarounds range from the inconvenient, to the absurd.

* Colleagues or line managers are required to proof read and correct the text
* Text can be pre-checked and pasted into the field once complete
* Users can goodly enter they're text without no grammar or speelingz are being incorrect

### Running a Spellcheck on a Field

Whilst it's true that these fields cannot be readily spell-checked, a Word Macro can be triggered by an "on-exit" event to easily spell check a field in a "protected" Word 2003 document.

The following proof of concept illustrates this, along with a simple dialog to allow the user to skip the spellcheck if required. It is envisaged that the Macro would only be applied to fields with paragraph text, since checking name, address, date-of-birth, and other similar fields is likely to be counter productive.

The user sees a form as per the example below:

![empty form](https://github.com/jonathancraddock/protected-spellcheck/blob/master/prot-spell-1.png "Empty protected form.")

### Macro

A Macro, see below, is triggered automatically by the "Run Macro... On Exit" event.

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

The Macro can be applied to selected fields as shown below: 

![field dialog](https://github.com/jonathancraddock/protected-spellcheck/blob/master/prot-spell-2.png "Form field dialog.")

The user is prompted as they exit the field with the "Tab" key, or by clicking elsewhere in the document.
    
![prompt](https://github.com/jonathancraddock/protected-spellcheck/blob/master/prot-spell-3.png "Spell check prompt.")

Word's own spell-check is used to check the field spelling.

![suggestions](https://github.com/jonathancraddock/protected-spellcheck/blob/master/prot-spell-4.png "Suggestions.")

User sees an info dialog once the spell check is complete.

![complete](https://github.com/jonathancraddock/protected-spellcheck/blob/master/prot-spell-5.png "Spell check complete.")

And, the field is updated with any corrections that have taken place.

![corrected](https://github.com/jonathancraddock/protected-spellcheck/blob/master/prot-spell-6.png "Spelling corrected.")

At this point, the user continues with the form as usual.
