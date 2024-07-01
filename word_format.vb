Sub CAV_Makro()

'Formatvorlagen einkopieren
ActiveDocument.CopyStylesFromTemplate _
    Template:="Path/to/your/customStyle.docx"

'Nummerierte Überschriften formatieren. Bei "YourCustomHeading_num" die gewünschte nummerierte Überschrift nennen
For Each Paragraph In ActiveDocument.Paragraphs
    If Paragraph.style = ActiveDocument.Styles(wdStyleHeading1) And Paragraph.Range.ListFormat.ListType = wdListSimpleNumbering Then
        Paragraph.style = ActiveDocument.Styles("YourCustomHeading_num")
    ElseIf Paragraph.style = ActiveDocument.Styles(wdStyleHeading2) And Paragraph.Range.ListFormat.ListType = wdListSimpleNumbering Then
        Paragraph.style = ActiveDocument.Styles("YourCustomHeading_num")
    ElseIf Paragraph.style = ActiveDocument.Styles(wdStyleHeading3) And Paragraph.Range.ListFormat.ListType = wdListSimpleNumbering Then
        Paragraph.style = ActiveDocument.Styles("YourCustomHeading_num")
    ElseIf Paragraph.style = ActiveDocument.Styles(wdStyleHeading4) And Paragraph.Range.ListFormat.ListType = wdListSimpleNumbering Then
        Paragraph.style = ActiveDocument.Styles("YourCustomHeading_num")
    ElseIf Paragraph.style = ActiveDocument.Styles(wdStyleHeading5) And Paragraph.Range.ListFormat.ListType = wdListSimpleNumbering Then
        Paragraph.style = ActiveDocument.Styles("YourCustomHeading_num")
    ElseIf Paragraph.style = ActiveDocument.Styles(wdStyleHeading6) And Paragraph.Range.ListFormat.ListType = wdListSimpleNumbering Then
        Paragraph.style = ActiveDocument.Styles("YourCustomHeading_num")
    End If
Next Paragraph

'Überschriften formatieren. Bei "YourCustomHeading" die gewünschte Überschrift nennen
For Each Paragraph In ActiveDocument.Paragraphs
    If Paragraph.style = ActiveDocument.Styles(wdStyleHeading1) Then
        Paragraph.style = ActiveDocument.Styles("YourCustomHeading")
    ElseIf Paragraph.style = ActiveDocument.Styles(wdStyleHeading2) Then
        Paragraph.style = ActiveDocument.Styles("YourCustomHeading")
    ElseIf Paragraph.style = ActiveDocument.Styles(wdStyleHeading3) Then
        Paragraph.style = ActiveDocument.Styles("YourCustomHeading")
    ElseIf Paragraph.style = ActiveDocument.Styles(wdStyleHeading4) Then
        Paragraph.style = ActiveDocument.Styles("YourCustomHeading")
    ElseIf Paragraph.style = ActiveDocument.Styles(wdStyleHeading5) Then
        Paragraph.style = ActiveDocument.Styles("YourCustomHeading")
    ElseIf Paragraph.style = ActiveDocument.Styles(wdStyleHeading6) Then
        Paragraph.style = ActiveDocument.Styles("YourCustomHeading")
    End If
Next Paragraph

'FLIEßTEXT

'Fließtext formatieren. Formatiert auch Paragraphen mit Border/Kasten sowie Aufzählungen in Kästen.
For Each Paragraph In ActiveDocument.Paragraphs
    If Paragraph.style = ActiveDocument.Styles(wdStyleNormal) And Paragraph.Borders.Enable = True Then
        Paragraph.style = ActiveDocument.Styles("CustomParagraph Kasten")
    ElseIf Paragraph.Borders.Enable = True And Paragraph.style = ActiveDocument.Styles(wdStyleListParagraph) And Paragraph.Range.ListFormat.ListType = wdListBullet And Paragraph.Range.ListFormat.ListLevelNumber = 1 Then
        Paragraph.style = ActiveDocument.Styles("CustomParagraph Kasten-Aufzählung")
    ElseIf Paragraph.Borders.Enable = True And Paragraph.style = ActiveDocument.Styles(wdStyleListParagraph) And Paragraph.Range.ListFormat.ListType = wdListBullet And Paragraph.Range.ListFormat.ListLevelNumber = 2 Then
        Paragraph.style = ActiveDocument.Styles("CustomParagraph Kasten-Aufzählung")
        Paragraph.Range.ListFormat.ListLevelNumber = 2
        Paragraph.TabStops.Add Position:=36, Alignment:=wdAlignTabLeft, Leader:=wdTabLeaderSpaces
    ElseIf Paragraph.Borders.Enable = True And Paragraph.style = ActiveDocument.Styles(wdStyleListParagraph) And Paragraph.Range.ListFormat.ListType = wdListBullet And Paragraph.Range.ListFormat.ListLevelNumber = 3 Then
        Paragraph.style = ActiveDocument.Styles("CustomParagraph Kasten-Aufzählung")
        Paragraph.Range.ListFormat.ListLevelNumber = 3
        Paragraph.TabStops.Add Position:=72, Alignment:=wdAlignTabLeft, Leader:=wdTabLeaderSpaces
    ElseIf Paragraph.Borders.Enable = True And Paragraph.style = ActiveDocument.Styles(wdStyleListParagraph) And Paragraph.Range.ListFormat.ListType = wdListBullet And Paragraph.Range.ListFormat.ListLevelNumber = 4 Then
        Paragraph.style = ActiveDocument.Styles("CustomParagraph Kasten-Aufzählung")
        Paragraph.Range.ListFormat.ListLevelNumber = 4
        Paragraph.TabStops.Add Position:=108, Alignment:=wdAlignTabLeft, Leader:=wdTabLeaderSpaces
    ElseIf Paragraph.Borders.Enable = True And Paragraph.style = ActiveDocument.Styles(wdStyleListParagraph) And Paragraph.Range.ListFormat.ListType = wdListSimpleNumbering And Paragraph.Range.ListFormat.ListLevelNumber = 1 Then
        Paragraph.style = ActiveDocument.Styles("CustomParagraph Kasten-Aufzählung")
    ElseIf Paragraph.Borders.Enable = True And Paragraph.style = ActiveDocument.Styles(wdStyleListParagraph) And Paragraph.Range.ListFormat.ListType = wdListSimpleNumbering And Paragraph.Range.ListFormat.ListLevelNumber = 2 Then
        Paragraph.style = ActiveDocument.Styles("CustomParagraph Kasten-Aufzählung")
        Paragraph.Range.ListFormat.ListLevelNumber = 2
        Paragraph.TabStops.Add Position:=36, Alignment:=wdAlignTabLeft, Leader:=wdTabLeaderSpaces
    ElseIf Paragraph.Borders.Enable = True And Paragraph.style = ActiveDocument.Styles(wdStyleListParagraph) And Paragraph.Range.ListFormat.ListType = wdListSimpleNumbering And Paragraph.Range.ListFormat.ListLevelNumber = 3 Then
        Paragraph.style = ActiveDocument.Styles("CustomParagraph Kasten-Aufzählung")
        Paragraph.Range.ListFormat.ListLevelNumber = 3
        Paragraph.TabStops.Add Position:=72, Alignment:=wdAlignTabLeft, Leader:=wdTabLeaderSpaces
    ElseIf Paragraph.Borders.Enable = True And Paragraph.style = ActiveDocument.Styles(wdStyleListParagraph) And Paragraph.Range.ListFormat.ListType = wdListSimpleNumbering And Paragraph.Range.ListFormat.ListLevelNumber = 4 Then
        Paragraph.style = ActiveDocument.Styles("CustomParagraph Kasten-Aufzählung")
        Paragraph.Range.ListFormat.ListLevelNumber = 4
        Paragraph.TabStops.Add Position:=108, Alignment:=wdAlignTabLeft, Leader:=wdTabLeaderSpaces
    ElseIf Paragraph.style = ActiveDocument.Styles(wdStyleNormal) And Paragraph.Borders.Enable = False Then
        Paragraph.style = ActiveDocument.Styles("CustomFließtext")
    End If
Next Paragraph

'AUFZÄHLUNGEN
'Nummerierte Listen. Bei "YourCustomList_num" die gewünschte nummerierte Aufzählung nennen
For Each Paragraph In ActiveDocument.Paragraphs
    If Paragraph.style = ActiveDocument.Styles(wdStyleListParagraph) And Paragraph.Range.ListFormat.ListType = wdListSimpleNumbering And Paragraph.Range.ListFormat.ListLevelNumber = 1 Then
        Paragraph.style = ActiveDocument.Styles("YourCustomList_num")
    ElseIf Paragraph.style = ActiveDocument.Styles(wdStyleListParagraph) And Paragraph.Range.ListFormat.ListType = wdListSimpleNumbering And Paragraph.Range.ListFormat.ListLevelNumber = 2 Then
        Paragraph.style = ActiveDocument.Styles("YourCustomList_num")
        Paragraph.Range.ListFormat.ListLevelNumber = 2
        Paragraph.TabStops.Add Position:=36, Alignment:=wdAlignTabLeft, Leader:=wdTabLeaderSpaces
    ElseIf Paragraph.style = ActiveDocument.Styles(wdStyleListParagraph) And Paragraph.Range.ListFormat.ListType = wdListSimpleNumbering And Paragraph.Range.ListFormat.ListLevelNumber = 3 Then
        Paragraph.style = ActiveDocument.Styles("YourCustomList_num")
        Paragraph.Range.ListFormat.ListLevelNumber = 3
        Paragraph.TabStops.Add Position:=72, Alignment:=wdAlignTabLeft, Leader:=wdTabLeaderSpaces
    ElseIf Paragraph.style = ActiveDocument.Styles(wdStyleListParagraph) And Paragraph.Range.ListFormat.ListType = wdListSimpleNumbering And Paragraph.Range.ListFormat.ListLevelNumber = 4 Then
        Paragraph.style = ActiveDocument.Styles("YourCustomList_num")
        Paragraph.Range.ListFormat.ListLevelNumber = 4
        Paragraph.TabStops.Add Position:=108, Alignment:=wdAlignTabLeft, Leader:=wdTabLeaderSpaces
    End If
Next

'Listen mit Bullets. Bei "YourCustomList_Bullet" die gewünschte Bullet Aufzählung nennen
For Each Paragraph In ActiveDocument.Paragraphs
    If Paragraph.style = ActiveDocument.Styles(wdStyleListParagraph) And Paragraph.Range.ListFormat.ListType = wdListBullet And Paragraph.Range.ListFormat.ListLevelNumber = 1 Then
        Paragraph.style = ActiveDocument.Styles("YourCustomList_Bullet")
    ElseIf Paragraph.style = ActiveDocument.Styles(wdStyleListParagraph) And Paragraph.Range.ListFormat.ListType = wdListBullet And Paragraph.Range.ListFormat.ListLevelNumber = 2 Then
        Paragraph.style = ActiveDocument.Styles("YourCustomList_Bullet")
        Paragraph.Range.ListFormat.ListLevelNumber = 2
        Paragraph.TabStops.Add Position:=36, Alignment:=wdAlignTabLeft, Leader:=wdTabLeaderSpaces
    ElseIf Paragraph.style = ActiveDocument.Styles(wdStyleListParagraph) And Paragraph.Range.ListFormat.ListType = wdListBullet And Paragraph.Range.ListFormat.ListLevelNumber = 3 Then
        Paragraph.style = ActiveDocument.Styles("YourCustomList_Bullet")
        Paragraph.Range.ListFormat.ListLevelNumber = 3
        Paragraph.TabStops.Add Position:=72, Alignment:=wdAlignTabLeft, Leader:=wdTabLeaderSpaces
    ElseIf Paragraph.style = ActiveDocument.Styles(wdStyleListParagraph) And Paragraph.Range.ListFormat.ListType = wdListBullet And Paragraph.Range.ListFormat.ListLevelNumber = 4 Then
        Paragraph.style = ActiveDocument.Styles("YourCustomList_Bullet")
        Paragraph.Range.ListFormat.ListLevelNumber = 4
        Paragraph.TabStops.Add Position:=108, Alignment:=wdAlignTabLeft, Leader:=wdTabLeaderSpaces
    End If
Next

'Listen mit Alphabets. Bei "YourCustomList_alphabet" die gewünschte alphabetische Aufzählung nennen
For Each Paragraph In ActiveDocument.Paragraphs
    If Paragraph.style = ActiveDocument.Styles(wdStyleListParagraph) And Paragraph.Range.ListFormat.ListType = wdListSimpleNumbering And Paragraph.Range.ListFormat.ListLevelNumber = 1 Then
        Paragraph.style = ActiveDocument.Styles("YourCustomList_alphabet")
    ElseIf Paragraph.style = ActiveDocument.Styles(wdStyleListParagraph) And Paragraph.Range.ListFormat.ListType = wdListSimpleNumbering And Paragraph.Range.ListFormat.ListLevelNumber = 2 Then
        Paragraph.style = ActiveDocument.Styles("YourCustomList_alphabet")
        Paragraph.Range.ListFormat.ListLevelNumber = 2
        Paragraph.TabStops.Add Position:=36, Alignment:=wdAlignTabLeft, Leader:=wdTabLeaderSpaces
    ElseIf Paragraph.style = ActiveDocument.Styles(wdStyleListParagraph) And Paragraph.Range.ListFormat.ListType = wdListSimpleNumbering And Paragraph.Range.ListFormat.ListLevelNumber = 3 Then
        Paragraph.style = ActiveDocument.Styles("YourCustomList_alphabet")
        Paragraph.Range.ListFormat.ListLevelNumber = 3
        Paragraph.TabStops.Add Position:=72, Alignment:=wdAlignTabLeft, Leader:=wdTabLeaderSpaces
    ElseIf Paragraph.style = ActiveDocument.Styles(wdStyleListParagraph) And Paragraph.Range.ListFormat.ListType = wdListSimpleNumbering And Paragraph.Range.ListFormat.ListLevelNumber = 4 Then
        Paragraph.style = ActiveDocument.Styles("YourCustomList_alphabet")
        Paragraph.Range.ListFormat.ListLevelNumber = 4
        Paragraph.TabStops.Add Position:=108, Alignment:=wdAlignTabLeft, Leader:=wdTabLeaderSpaces
    End If
Next

'Listen mit Spiegelstrichen. Bei "YourCustomList_line" die gewünschte Spiegelstrich-Aufzählung nennen
For Each Paragraph In ActiveDocument.Paragraphs
    If Paragraph.style = ActiveDocument.Styles(wdStyleListParagraph) And Paragraph.Range.ListFormat.ListType = wdListBullet And Paragraph.Range.ListFormat.ListLevelNumber = 1 Then
        Paragraph.style = ActiveDocument.Styles("YourCustomList_line")
    ElseIf Paragraph.style = ActiveDocument.Styles(wdStyleListParagraph) And Paragraph.Range.ListFormat.ListType = wdListBullet And Paragraph.Range.ListFormat.ListLevelNumber = 2 Then
        Paragraph.style = ActiveDocument.Styles("YourCustomList_line")
        Paragraph.Range.ListFormat.ListLevelNumber = 2
        Paragraph.TabStops.Add Position:=36, Alignment:=wdAlignTabLeft, Leader:=wdTabLeaderSpaces
    ElseIf Paragraph.style = ActiveDocument.Styles(wdStyleListParagraph) And Paragraph.Range.ListFormat.ListType = wdListBullet And Paragraph.Range.ListFormat.ListLevelNumber = 3 Then
        Paragraph.style = ActiveDocument.Styles("YourCustomList_line")
        Paragraph.Range.ListFormat.ListLevelNumber = 3
        Paragraph.TabStops.Add Position:=72, Alignment:=wdAlignTabLeft, Leader:=wdTabLeaderSpaces
    ElseIf Paragraph.style = ActiveDocument.Styles(wdStyleListParagraph) And Paragraph.Range.ListFormat.ListType = wdListBullet And Paragraph.Range.ListFormat.ListLevelNumber = 4 Then
        Paragraph.style = ActiveDocument.Styles("YourCustomList_line")
        Paragraph.Range.ListFormat.ListLevelNumber = 4
        Paragraph.TabStops.Add Position:=108, Alignment:=wdAlignTabLeft, Leader:=wdTabLeaderSpaces
    End If
Next

'ABSATZFORMATIERUNGEN
'Zitate
For Each Paragraph In ActiveDocument.Paragraphs
    If Paragraph.style = ActiveDocument.Styles(wdStyleQuote) Then
        Paragraph.style = ActiveDocument.Styles("YourCustomStyle")
    End If
Next

'Fußnote formatieren. Formatiert die Absatzformatierung von Fußnotentext
For Each Footnote In ActiveDocument.Footnotes
    Footnote.Range.style = ActiveDocument.Styles("YourCustomStyle")
Next

'KOPF- UND FUSSZEILEN
'Kopfzeile entfernen
ActiveDocument.Sections(1).Headers(wdHeaderFooterPrimary).Range.Delete

'Fußzeile entfernen
ActiveDocument.Sections(1).Footers(wdHeaderFooterPrimary).Range.Delete

'Fußzeile einfügen
ActiveDocument.Sections(1).Footers(wdHeaderFooterPrimary).PageNumbers.Add _
    PageNumberAlignment:=wdAlignPageNumberRight, _
    FirstPage:=True
ActiveDocument.Sections(1).Footers(wdHeaderFooterPrimary).Range.InsertAfter Format(Date, "dd.mm.yyyy")
ActiveDocument.Sections(1).Footers(wdHeaderFooterPrimary).Range.InsertAfter " - " & ActiveDocument.Name


'ZEICHENFORMATIERUNGEN (formatiert keine Absätze, sondern nur Zeichen)
Dim Word As Range

'kursiv. Formatiert alle kursiven Zeichen 
For Each Paragraph In ActiveDocument.Paragraphs
    For Each Word In Paragraph.Range.Words
        If Word.Font.Italic = True Then
            Word.style = ActiveDocument.Styles("YourCustomStyle")
        End If
    Next Word
Next Paragraph

'fett. Die Custom Heading Styles müssen explizit ausgeschlossen werden.
For Each Paragraph In ActiveDocument.Paragraphs
    For Each Word In Paragraph.Range.Words
        If Word.Font.Bold = True And Paragraph.style <> ActiveDocument.Styles("YourCustomHeadingStyles") _
        And Paragraph.style <> ActiveDocument.Styles("YourCustomHeadingStyles") _
        And Paragraph.style <> ActiveDocument.Styles("YourCustomHeadingStyles") _
        And Paragraph.style <> ActiveDocument.Styles("YourCustomHeadingStyles") _
        And Paragraph.style <> ActiveDocument.Styles("YourCustomHeadingStyles") _
        And Paragraph.style <> ActiveDocument.Styles("YourCustomHeadingStyles") _
        And Paragraph.style <> ActiveDocument.Styles("YourCustomHeadingStyles num.") _
        And Paragraph.style <> ActiveDocument.Styles("YourCustomHeadingStyles num.") _
        And Paragraph.style <> ActiveDocument.Styles("YourCustomHeadingStyles num.") _
        And Paragraph.style <> ActiveDocument.Styles("YourCustomHeadingStyles num.") _
        And Paragraph.style <> ActiveDocument.Styles("YourCustomHeadingStyles num.") _
        And Paragraph.style <> ActiveDocument.Styles("YourCustomHeadingStyles num.") Then
            Word.style = ActiveDocument.Styles("YourCustomStyle")
        End If
    Next
Next

'EINZELZEICHEN
'Doppelte Leerzeichen entfernen
With ActiveDocument.Content.Find
    .ClearFormatting
    .Text = "  "
    .Replacement.Text = " "
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = False
    .MatchWholeWord = False
    .MatchByte = False
    .MatchAllWordForms = False
    .MatchSoundsLike = False
    .MatchWildcards = False
    .MatchFuzzy = False
    .Execute Replace:=wdReplaceAll
End With

'Doppelte Absatzeichen entfernen
With ActiveDocument.Content.Find
    .ClearFormatting
    .Text = "^p^p"
    .Replacement.Text = "^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = False
    .MatchWholeWord = False
    .MatchByte = False
    .MatchAllWordForms = False
    .MatchSoundsLike = False
    .MatchWildcards = False
    .MatchFuzzy = False
    .Execute Replace:=wdReplaceAll
End With

'ANFÜHRUNGSZEICHEN
'linke doppelte Anführungszeichen
Dim leftdoublequotes() As Variant
    leftdoublequotes = Array(Chr(34), ChrW(&H201C), ChrW(&H201E))
    
    Dim i As Long

    ' Loop through each type of left double quotation mark and replace it
    For i = LBound(leftdoublequotes) To UBound(leftdoublequotes)
        With ActiveDocument.Content.Find
            .ClearFormatting
            .Text = " " & leftdoublequotes(i)
            .Replacement.Text = " " & ChrW(&H00BB) ' Guillimets
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = False
            .MatchWholeWord = False
            .MatchByte = False
            .MatchAllWordForms = False
            .MatchSoundsLike = False
            .MatchWildcards = False
            .MatchFuzzy = False
            .Execute Replace:=wdReplaceAll
        End With
    Next i

'rechte doppelte Anführungszeichen
Dim rightdoublequotes() As Variant
    rightdoublequotes = Array(Chr(34), ChrW(&H201D), ChrW(&H2033))
    
    ' Loop through each type of right double quotation mark and replace it
    For i = LBound(rightdoublequotes) To UBound(rightdoublequotes)
        With ActiveDocument.Content.Find
            .ClearFormatting
            .Text = rightdoublequotes(i) & " "
            .Replacement.Text = ChrW(&H00AB) & " " ' Guillimets
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = False
            .MatchWholeWord = False
            .MatchByte = False
            .MatchAllWordForms = False
            .MatchSoundsLike = False
            .MatchWildcards = False
            .MatchFuzzy = False
            .Execute Replace:=wdReplaceAll
        End With
    Next i

'linke doppelte Anführungszeichen MIT ABSATZ
    ' Loop through each type of left double quotation mark and replace it
    For i = LBound(leftdoublequotes) To UBound(leftdoublequotes)
        With ActiveDocument.Content.Find
            .ClearFormatting
            .Text = "^p" & leftdoublequotes(i)
            .Replacement.Text = "^p" & ChrW(&H00BB) ' Guillimets
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = False
            .MatchWholeWord = False
            .MatchByte = False
            .MatchAllWordForms = False
            .MatchSoundsLike = False
            .MatchWildcards = False
            .MatchFuzzy = False
            .Execute Replace:=wdReplaceAll
        End With
    Next i

'rechte doppelte Anführungszeichen MIT ABSATZ
    ' Loop through each type of right double quotation mark and replace it
    For i = LBound(rightdoublequotes) To UBound(rightdoublequotes)
        With ActiveDocument.Content.Find
            .ClearFormatting
            .Text = rightdoublequotes(i) & "^p"
            .Replacement.Text = ChrW(&H00AB) & "^p" ' Guillimets
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = False
            .MatchWholeWord = False
            .MatchByte = False
            .MatchAllWordForms = False
            .MatchSoundsLike = False
            .MatchWildcards = False
            .MatchFuzzy = False
            .Execute Replace:=wdReplaceAll
        End With
    Next i


'linke einfache Anführungszeichen
Dim leftsinglequotes() As Variant
    leftsinglequotes = Array(Chr(39), ChrW(&H2018), ChrW(&H201A))
    
    ' Loop through each type of left single quotation mark and replace it
    For i = LBound(leftsinglequotes) To UBound(leftsinglequotes)
        With ActiveDocument.Content.Find
            .ClearFormatting
            .Text = " " & leftsinglequotes(i)
            .Replacement.Text = " " & ChrW(&H203A)
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = False
            .MatchWholeWord = False
            .MatchByte = False
            .MatchAllWordForms = False
            .MatchSoundsLike = False
            .MatchWildcards = False
            .MatchFuzzy = False
            .Execute Replace:=wdReplaceAll
        End With
    Next i

'rechte einfache Anführungszeichen
Dim rightsinglequotes() As Variant
    rightsinglequotes = Array(Chr(39), ChrW(&H2019), ChrW(&H2032))
    
    ' Loop through each type of right single quotation mark and replace it
    For i = LBound(rightsinglequotes) To UBound(rightsinglequotes)
        With ActiveDocument.Content.Find
            .ClearFormatting
            .Text = rightsinglequotes(i) & " "
            .Replacement.Text = ChrW(&H2039) & " "
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = False
            .MatchWholeWord = False
            .MatchByte = False
            .MatchAllWordForms = False
            .MatchSoundsLike = False
            .MatchWildcards = False
            .MatchFuzzy = False
            .Execute Replace:=wdReplaceAll
        End With
    Next i

'Gedankenstriche
With ActiveDocument.Content.Find
    .ClearFormatting
    .Text = " - "
    .Replacement.Text = " – "
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = False
    .MatchWholeWord = False
    .MatchByte = False
    .MatchAllWordForms = False
    .MatchSoundsLike = False
    .MatchWildcards = False
    .MatchFuzzy = False
    .Execute Replace:=wdReplaceAll
End With

'Striche zwischen zwei Zahlen
With ActiveDocument.Content.Find
    .ClearFormatting
    .Text = "([0-9])-([0-9])"
    .Replacement.Text = "\1–\2"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = False
    .MatchWholeWord = False
    .MatchByte = False
    .MatchAllWordForms = False
    .MatchSoundsLike = False
    .MatchWildcards = True
    .MatchFuzzy = False
    .Execute Replace:=wdReplaceAll
End With

'Auslassungszeichen
With ActiveDocument.Content.Find
    .ClearFormatting
    .Text = "(…)"
    .Replacement.Text = "[...]"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = False
    .MatchWholeWord = False
    .MatchByte = False
    .MatchAllWordForms = False
    .MatchSoundsLike = False
    .MatchWildcards = False
    .MatchFuzzy = False
    .Execute Replace:=wdReplaceAll
End With

'Inhaltsverzeichnis
    Selection.HomeKey Unit:=wdStory
    Selection.InsertNewPage
    Selection.HomeKey Unit:=wdStory
    Application.Templates( _
        "/Applications/Microsoft Word.app/Contents/Resources/de.lproj/Document Elements/Table of Contents.dotx" _ 'wenn auf einem anderen Rechner, dann muss Pfad neu eingefügt werden (Makro aufnehmen, IHVZ einfügen, Pfad kopieren)
        ).BuildingBlockEntries("Einfach").Insert Where:=Selection.Range, RichText _
        :=True

'Geschütztes Leerzeichen bei Abkürzungen einfügen
    Dim doc As Document
    Dim rng As Range
    Dim oldAbbr As String
    Dim newAbbr As String

    ' Set the active document to operate on
    Set doc = ActiveDocument

    ' Define the abbreviations and their replacements
    Dim abbrs As Variant
    Dim replacements As Variant

    abbrs = Array("z.B.", "d.h.", "u.a.", "o.ä.", "i.d.R.", "z. B.", "d. h.", "u. a.", "o. ä.", "i. d. R.")
    replacements = Array("z." & ChrW(160) & "B.", "d." & ChrW(160) & "h.", "u." & ChrW(160) & "a.", "o." & ChrW(160) & "ä.", "i." & ChrW(160) & "d." & ChrW(160) & "R.", "z." & ChrW(160) & "B.", "d." & ChrW(160) & "h.", "u." & ChrW(160) & "a.", "o." & ChrW(160) & "ä.", "i." & ChrW(160) & "d." & ChrW(160) & "R.")

    ' Loop through each abbreviation and replace it in the document
    For i = LBound(abbrs) To UBound(abbrs)
        oldAbbr = abbrs(i)
        newAbbr = replacements(i)

        ' Create a range for the entire document
        Set rng = doc.Content

        ' Find and replace the abbreviation
        With rng.Find
            .Text = oldAbbr
            .Replacement.Text = newAbbr
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
            .Execute Replace:=wdReplaceAll
        End With
    Next i


End Sub
