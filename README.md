Sub PowerPointPresentation()
    Dim pptApp As Object
    Dim pptPres As Object
    Dim pptSlide As Object

    ' Create a new PowerPoint application
    Set pptApp = CreateObject("PowerPoint.Application")
    pptApp.Visible = True

    ' Add a new presentation
    Set pptPres = pptApp.Presentations.Add

    ' Add the first slide
    Set pptSlide = pptPres.Slides.Add(1, 1) ' 1 = ppLayoutTitle
    With pptSlide
        .Shapes(1).TextFrame.TextRange.Text = "رواية \"الحي اللاتيني\""
        .Shapes(2).TextFrame.TextRange.Text = "عرض تحليلي للأقسام 1 إلى 5 من القسم الثاني"
    End With

    ' Add the second slide
    Set pptSlide = pptPres.Slides.Add(2, 2) ' 2 = ppLayoutText
    With pptSlide
        .Shapes(1).TextFrame.TextRange.Text = "تتبع الأحداث حسب الأقسام"
        .Shapes(2).TextFrame.TextRange.Text = "- انتقال البطل إلى الحي اللاتيني في باريس\n- لقاء شخصيات فرنسية متنوعة\n- تعمّق العلاقة مع جانين\n- تأمل البطل في هويته الذاتية"
    End With

    ' Add the third slide
    Set pptSlide = pptPres.Slides.Add(3, 2) ' 2 = ppLayoutText
    With pptSlide
        .Shapes(1).TextFrame.TextRange.Text = "دراسة الشخصيات"
        .Shapes(2).TextFrame.TextRange.Text = "- البطل: ذكي، حساس، يعيش صراعًا داخليًا\n- جانين: متحررة، رومانسية\n- الشخصيات الثانوية: تعكس تعددية التجربة الثقافية"
    End With

    ' Add the fourth slide
    Set pptSlide = pptPres.Slides.Add(4, 2) ' 2 = ppLayoutText
    With pptSlide
        .Shapes(1).TextFrame.TextRange.Text = "الفضاء السردي"
        .Shapes(2).TextFrame.TextRange.Text = "- المكان: الحي اللاتيني في باريس\n- الزمان: فترة ما بعد الحرب العالمية الثانية\n- الأجواء: التوتر والتأمل"
    End With

    ' Add the fifth slide
    Set pptSlide = pptPres.Slides.Add(5, 2) ' 2 = ppLayoutText
    With pptSlide
        .Shapes(1).TextFrame.TextRange.Text = "نمط السرد والرؤية السردية"
        .Shapes(2).TextFrame.TextRange.Text = "- نمط السرد: واقعي تأملي\n- الرؤية السردية: الرؤية من الداخل"
    End With

    ' Add the sixth slide
    Set pptSlide = pptPres.Slides.Add(6, 2) ' 2 = ppLayoutText
    With pptSlide
        .Shapes(1).TextFrame.TextRange.Text = "الأبعاد والمقاصد"
        .Shapes(2).TextFrame.TextRange.Text = "- الأبعاد: ثقافي، اجتماعي، نفسي\n- المقاصد: إبراز أهمية الهوية الثقافية، الدعوة إلى فهم الآخر"
    End With

    ' Save the presentation (optional)
    Dim filePath As String
    filePath = "C:\Users\Public\Documents\Presentation_Hay_Latin.pptx"
    pptPres.SaveAs filePath

    ' Release objects
    Set pptSlide = Nothing
    Set pptPres = Nothing
    Set pptApp = Nothing

    MsgBox "Presentation created and saved to: " & filePath
End Sub  
