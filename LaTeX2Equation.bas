Attribute VB_Name = "LaTeX2Equation"
Option Explicit

' API declaration for Sleep
#If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

' ==========================================
' Ribbon Callbacks
' ==========================================
Sub OnConvertSelectedText(control As IRibbonControl)
    ConvertSelectedText
End Sub

Sub OnConvertCurrentSlide(control As IRibbonControl)
    ConvertCurrentSlide
End Sub

Sub OnConvertAllSlides(control As IRibbonControl)
    ConvertAllSlides
End Sub

' ==========================================
' MACRO 1: Process only the currently selected text
' ==========================================
Sub ConvertSelectedText()
    If ActiveWindow.Selection.Type = ppSelectionText Then
        Call ProcessTextRange(ActiveWindow.Selection.TextRange)
        MsgBox "Selection processing complete!", vbInformation
    Else
        MsgBox "Please select some text first.", vbExclamation
    End If
End Sub

' ==========================================
' MACRO 2: Process all shapes on the current slide
' ==========================================
Sub ConvertCurrentSlide()
    Dim currentSlide As Slide
    Dim shp As Shape
    
    If ActiveWindow.ViewType <> ppViewNormal And ActiveWindow.ViewType <> ppViewSlide Then
        MsgBox "Please switch to Normal or Slide view.", vbExclamation
        Exit Sub
    End If
    
    Set currentSlide = ActiveWindow.View.Slide
    
    For Each shp In currentSlide.Shapes
        If shp.HasTextFrame Then
            If shp.TextFrame.HasText Then
                Call ProcessTextRange(shp.TextFrame.TextRange)
            End If
        End If
    Next shp
    
    MsgBox "Current slide processing complete!", vbInformation
End Sub

' ==========================================
' MACRO 3: Process the entire presentation
' ==========================================
Sub ConvertAllSlides()
    Dim sld As Slide
    Dim shp As Shape
    
    For Each sld In ActivePresentation.Slides
        ActiveWindow.View.GotoSlide sld.SlideIndex
        DoEvents
        Sleep 200
        
        For Each shp In sld.Shapes
            If shp.HasTextFrame Then
                If shp.TextFrame.HasText Then
                    Call ProcessTextRange(shp.TextFrame.TextRange)
                End If
            End If
        Next shp
    Next sld
    
    MsgBox "Entire presentation processing complete!", vbInformation
End Sub

' ==========================================
' CORE FUNCTION: Regex match and replace with equation
' ==========================================
Private Sub ProcessTextRange(ByVal oTxtRng As TextRange)
    Dim regEx As Object
    Dim matches As Object
    Dim i As Integer
    Dim m As Object
    Dim cleanText As String
    Dim startIdx As Long
    
    Set regEx = CreateObject("VBScript.RegExp")
    regEx.Global = True
    regEx.IgnoreCase = False
    regEx.Pattern = "(\${1,2})([\s\S]+?)(\${1,2})"
    
    Set matches = regEx.Execute(oTxtRng.Text)
    
    For i = matches.Count - 1 To 0 Step -1
        Set m = matches(i)
        cleanText = m.SubMatches(1)
        startIdx = m.FirstIndex + 1
        
        oTxtRng.Characters(startIdx, m.Length).Select
        DoEvents
        Sleep 100
        
        ActiveWindow.Selection.TextRange.Text = cleanText
        DoEvents
        Sleep 100
        
        oTxtRng.Characters(startIdx, Len(cleanText)).Select
        DoEvents
        Sleep 100
        
        Call ConvertSelectionToEquation
    Next i
End Sub

' ==========================================
' HELPER: Convert selected LaTeX to rendered equation
' ==========================================
Private Sub ConvertSelectionToEquation()
    On Error Resume Next
    
    Application.CommandBars.ExecuteMso "EquationInsertNew"
    DoEvents
    Sleep 1000
    
    Application.CommandBars.ExecuteMso "EquationProfessional"
    DoEvents
    Sleep 500
    
    SendKeys "{ESCAPE}", True
    DoEvents
    Sleep 300
    
    On Error GoTo 0
End Sub
