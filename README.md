### VBA! (aka the Worst Language)
visual basic for applications macros, word/excel
they are here because i cba to set up an actual website
feel free to send to other people 

#### Word

##### Verbatim scroll-to-top issue
Apparently this sometimes doesn't work for mac -- an alternative is to use [draft mode](https://www.dummies.com/software/microsoft-office/word/how-to-change-the-document-view-in-word-2016/#:~:text=The%20Draft%20view%20presents%20only%20basic%20text)


```
Private Sub Document_New()
    ActiveDocument.Windows(1).View.RevisionsFilter.Markup = wdRevisionsMarkupNone
End Sub

Private Sub Document_Open()
    ActiveDocument.Windows(1).View.RevisionsFilter.Markup = wdRevisionsMarkupNone
End Sub
```


#### Excel
¯\_(ツ)_/¯
