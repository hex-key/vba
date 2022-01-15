## VBA (aka the Worst Language)
small fixes for verbatim because shit hasnt been updated since 2013

### Scroll-to-top issue
Apparently this sometimes doesn't work for mac -- an alternative is to use **[draft mode](https://www.dummies.com/software/microsoft-office/word/how-to-change-the-document-view-in-word-2016/#:~:text=The%20Draft%20view%20presents%20only%20basic%20text)**

Instructions to install: 
- open a file with the verbatim template
- hit alt-F11 (windows) or opt-F11 (mac) or fn-opt-11 (mac with touch bar) to open the VBA editor
- on the top left there will be a file tree with a lot of icons, click the [+] sign to the left of the one that says Verbatim (Debate)
- click the [+] sign next to the folder called "Microsoft Word Objects" that shows up under it
- double-click on "ThisDocument" 
- you should be seeing a blank screen, paste in the code below and hit ctrl-S (windows) or command-S (mac) to save the changes 
- close the VBA editor and relaunch word
- hopefully no more random scrolling?

```
Private Sub Document_New()
    ActiveDocument.Windows(1).View.RevisionsFilter.Markup = wdRevisionsMarkupNone
End Sub

Private Sub Document_Open()
    ActiveDocument.Windows(1).View.RevisionsFilter.Markup = wdRevisionsMarkupNone
End Sub
```

### Cite box issues 
Should fix that thing when you paste your cites into the wiki and then they don't show up

NOTE: I probably haven't included every single invalid character. This is because my only source of knowledge on which characters are invalid is when I personally try to upload cites and they don't work. 
#### So, if you implement this and your cites still don't show up, [submitting an issue](https://github.com/hex-key/vba/issues/new) with the cite that isn't working would be  super helpful for me to continue updating this!

Instructions to install: 
- open a file with the verbatim template
- hit alt-F11 (windows) or opt-F11 (mac) or fn-opt-11 (mac with touch bar) to open the VBA editor
- on the top left there will be a file tree with a lot of icons, click the [+] sign to the left of the one that says Verbatim (Debate)
- click the [+] sign next to the folder called "Modules" that shows up under it
- double-click on "Caselist"
- select all and delete everything in the file 
- copy and paste the code from **[Caselist.bas](https://github.com/hex-key/vba/blob/main/Caselist.bas)**
- you should be seeing a blank screen, paste in the code below and hit ctrl-S (windows) or command-S (mac) to save the changes 
- close the VBA editor and relaunch word
- hopefully never have to type "cites not working see os" again
