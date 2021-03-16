' Start Lodel
' Generates a menu to handle automatically OpenEdition templates and macros

Private Function getWordLang() As String
    ' https://msdn.microsoft.com/en-us/library/aa432635%28v=office.12%29.aspx
    Select Case Application.Language
        Case msoLanguageIDEnglishAUS, msoLanguageIDEnglishBelize, msoLanguageIDEnglishCanadian, msoLanguageIDEnglishCaribbean, msoLanguageIDEnglishIndonesia, msoLanguageIDEnglishIreland, msoLanguageIDEnglishJamaica, msoLanguageIDEnglishNewZealand, msoLanguageIDEnglishPhilippines, msoLanguageIDEnglishSouthAfrica, msoLanguageIDEnglishTrinidadTobago, msoLanguageIDEnglishUK, msoLanguageIDEnglishUS, msoLanguageIDEnglishZimbabwe
            getWordLang = "en"
        Case msoLanguageIDFrench, msoLanguageIDFrenchCameroon, msoLanguageIDFrenchCanadian, msoLanguageIDFrenchCotedIvoire, msoLanguageIDFrenchHaiti, msoLanguageIDFrenchLuxembourg, msoLanguageIDFrenchMali, msoLanguageIDFrenchMonaco, msoLanguageIDFrenchMorocco, msoLanguageIDFrenchReunion, msoLanguageIDFrenchSenegal, msoLanguageIDFrenchWestIndies, msoLanguageIDFranchCongoDRC, msoLanguageIDBelgianFrench, msoLanguageIDSwissFrench
            getWordLang = "fr"
        Case Else
            getWordLang = "en" ' default
    End Select
End Function

' Testing OS: http://www.rondebruin.nl/mac/mac001.htm
Private Function getOs() As String
    ' TODO: A tester sur mac (+ les paths)
    #If Mac Then
        getOs = "mac"
    #Else
        getOs = "win"
    #End If
End Function

Private Function trad(id As String, Optional lang As String = "")
    Dim key As String
    If lang = "" Then
        lang = getWordLang()
    End If
    key = lang + "." + id
    Select Case key
        Case "en.start"
            trad = "Start styling for Lodel"
        Case "fr.start"
            trad = "Démarrer le stylage pour Lodel"
        Case Else
            trad = "undefined"
    End Select
End Function

Private Sub generateStartLodelMenu()
    Dim menuBar As CommandBar
    Dim subMenu As CommandBarControl
    Dim subSubMenu As CommandBarControl
    Dim menuItem As CommandBarControl
    Dim subMenuItem As CommandBarControl

    Set menuBar = CommandBars.Add(menuBar:=False, Position:=msoBarTop, Name:="Start Lodel", Temporary:=True)
    menuBar.Visible = True

    Set menuItem = menuBar.Controls.Add(Type:=msoControlButton)
    With menuItem
        .Caption = trad("start")
        .OnAction = "startRevuesOrgDefault"
        .Style = msoButtonCaption
    End With
End Sub

Sub doStart(tpl As String)
    Dim macro As String
    Dim os As String
    ' Fix pour gerer les floats correctement dans toutes les langues. Voir : http://stackoverflow.com/questions/16191557/vba-word-changing-decimal-separator
    Dim strDecimal As String
    strDecimal = Application.International(wdDecimalSeparator)
    os = getOs()
    macro = Options.DefaultFilePath(Path:=wdUserTemplatesPath) + "\oe_macros_" + os + ".dotm"

    If ActiveWindow.View.SplitSpecial = wdPaneNone Then
        ActiveWindow.ActivePane.View.Type = wdNormalView
    Else
        ActiveWindow.View.Type = wdNormalView
    End If
    ActiveWindow.StyleAreaWidth = CentimetersToPoints(3)
    ActiveDocument.ActiveWindow.View.ShowAll = True
    ActiveDocument.FormattingShowFont = True
    ActiveDocument.FormattingShowParagraph = True
    ActiveDocument.FormattingShowNumbering = True
    ActiveDocument.FormattingShowFilter = wdShowFilterStylesInUse
    AddIns.Add FileName:=macro, Install:=True
    ActiveDocument.UpdateStylesOnOpen = True
    ActiveDocument.AttachedTemplate = tpl
    ActiveWindow.View.ShowBookmarks = True
    ' Afficher les noms de substitution quand on change le nom d'un style natif (Word 2007 et supérieurs uniquement)
    If CDbl(Replace(Application.Version, ".", strDecimal)) > 11 Then ActiveDocument.FormattingShowUserStyleName = True
End Sub

Sub startRevuesOrgDefault()
    Dim tpl As String
    Dim wordLang As String
    wordLang = getWordLang()
    If wordLang = "fr" Then
        tpl = Options.DefaultFilePath(Path:=wdUserTemplatesPath) + "\oe_modele_fr.dotm"
    Else
        tpl = Options.DefaultFilePath(Path:=wdUserTemplatesPath) + "\oe_template_en.dotm"
    End If
    doStart tpl
End Sub


Sub AutoExec()
    Call generateStartLodelMenu
End Sub

