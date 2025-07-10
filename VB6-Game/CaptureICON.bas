Attribute VB_Name = "CaptureICON"
Public Sub CaptureICO(Source_FORM_or_PICBOX, Dest_FORM As Form, Dest_PICBOX As PictureBox, Dest_IMAGELIST As ImageList, SaveInImageList As Boolean, X As Integer, Y As Integer, Width As Integer, Height As Integer)
    Dim long_side As Integer
    Dim Last_width As Integer, Last_height As Integer
    Dim Last_Image_Index As Integer
    Dim Last_AutoRedraw_of_Dest As Boolean
    Dim Last_ScaleMode_of_Dest As Integer
    Dim Last_Visible_of_Dest As Boolean
    Dim Last_ScaleMode_of_Dest_Form As Integer
    
    ' Get user Settings
    Last_AutoRedraw_of_Dest = Dest_PICBOX.AutoRedraw
    Last_ScaleMode_of_Dest = Dest_PICBOX.ScaleMode
    Last_Visible_of_Dest = Dest_PICBOX.Visible
    Last_ScaleMode_of_Dest_Form = Dest_FORM.ScaleMode
    
    ' Set Defaults
    Dest_PICBOX.AutoRedraw = True
    Dest_PICBOX.ScaleMode = 3
    Dest_PICBOX.Visible = False
    Dest_FORM.ScaleMode = 3
    
    ' Clear PictureBox
    Dest_PICBOX.Picture = Nothing
    
    ' Get sides of dest
    Last_width = Dest_PICBOX.Width
    Last_height = Dest_PICBOX.Height
    
    ' Get Largest Side of image
    If Height >= Width Then long_side = Height
    If Height <= Width Then long_side = Width
    
    ' Get Last Index in ImageList
    Last_Image_Index = Dest_IMAGELIST.ListImages.Count + 1
    
    ' Temporary resize the dest.
    With Dest_PICBOX
        .Height = long_side
        .Width = long_side
    End With
    
    ' Get the Image From Source
    retval = BitBlt&(Dest_PICBOX.hdc, 0, 0, Width, Height, Source_FORM_or_PICBOX.hdc, X, Y, SRCCOPY)
        
    ' Place the image to the ImageList
    Call Dest_IMAGELIST.ListImages.Add(Last_Image_Index, "", Dest_PICBOX.Image)
    
    ' Set Transparent Color
    Dest_IMAGELIST.MaskColor = &H8000000F
    
    ' Extract as icon
    Dest_PICBOX.Picture = Dest_IMAGELIST.ListImages(Last_Image_Index).ExtractIcon
    
    ' Clear Image (if user want!)
    If Not SaveInImageList Then Dest_IMAGELIST.ListImages.Remove (Last_Image_Index)
    
    ' Set the org. Sides of the Dest.
    Dest_PICBOX.Width = Last_width
    Dest_PICBOX.Height = Last_height
    
    ' Set user Settings
    Dest_PICBOX.AutoRedraw = Last_AutoRedraw_of_Dest
    Dest_PICBOX.ScaleMode = Last_ScaleMode_of_Dest
    Dest_PICBOX.Visible = Last_Visible_of_Dest
    Dest_FORM.ScaleMode = Last_ScaleMode_of_Dest_Form
End Sub



