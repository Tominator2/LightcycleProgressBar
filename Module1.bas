Sub InsertSmallLightcycles()

  ' Insert Progress Lightcycles on slides
  
  ' This expects the 2 image to be in same directory as the presentation!
  Dim picCycle As String
  picCycle = "TRON-sml.png"
          
  Dim picTrail As String
  picTrail = "trail-20x40.png"
          
  Dim sldHeight As Integer
  sldHeight = ActivePresentation.PageSetup.SlideHeight
  
  Dim sldWidth As Integer
  sldWidth = ActivePresentation.PageSetup.SlideWidth
  
  Dim totalSlides As Integer
  totalSlides = ActivePresentation.Slides.Count
  
  Dim sldNo As Integer
  Dim picX As Integer
  
  Dim sld As Slide
  
  ' Loop Through Each Slide
  For Each sld In ActivePresentation.Slides
  
    sld.Select
  
    Dim img As Shape
  
    ' Remove any existing lightcycles or trails (based on the top of the
    ' image position)
    For Each img In sld.Shapes
      If img.Top = sldHeight - 20 Then
          img.Delete
      End If
    Next img

    sldNo = sld.SlideNumber

    If sldNo = 1 Then
        ' First slide
        picX = 0
    ElseIf sldNo = totalSlides Then
        ' Last slide
        picX = sldWidth
    Else
        ' Other slides
        picX = ((sldNo - 1) / (totalSlides - 1)) * sldWidth
    End If
    
    ' Add trail if not on the first slide
    If sldNo <> 1 Then
      sld.Shapes.AddPicture(FileName:=picTrail, LinkToFile:=msoFalse, SaveWithDocument:=msoTrue, Left:=0, Top:=(sldHeight - 20), Width:=(picX - 25), Height:=20).Select
    End If

    ' Add cycle
    sld.Shapes.AddPicture(FileName:=picCycle, LinkToFile:=msoFalse, SaveWithDocument:=msoTrue, Left:=(picX - 35), Top:=(sldHeight - 20), Width:=86, Height:=20).Select
    
  Next sld
  
End Sub

