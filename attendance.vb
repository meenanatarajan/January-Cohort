Sub ATTENDANCE()

Dim k As Integer
Dim i As Integer
Dim j As Integer
  
  For i = 2 To 200
           For j = 2 To 128
       
               If (ActiveWorkbook.Worksheets("Zoom 1").Cells(i, 1).Value = Cells(j, 1)) Then
         Cells(j, 4).Value = "Present"
         ActiveWorkbook.Worksheets("Zoom 1").Cells(i, 10).Value = " Valid"
                  If (IsEmpty(Cells(j, 6).Value)) Then
            Cells(j, 6) = ActiveWorkbook.Worksheets("Zoom 1").Cells(i, 8).Value
                 End If
                Cells(j, 6) = ActiveWorkbook.Worksheets("Zoom 1").Cells(i, 8).Value
           Cells(j, 7) = Cells(j, 7) + ActiveWorkbook.Worksheets("Zoom 1").Cells(i, 5).Value
   
            End If
                 
     Next j
 Next i
 
 
   For i = 2 To 200
           For j = 2 To 128
       
               If (ActiveWorkbook.Worksheets("Zoom 2").Cells(i, 1).Value = Cells(j, 1)) Then
         Cells(j, 4).Value = "Present"
          ActiveWorkbook.Worksheets("Zoom 2").Cells(i, 10).Value = " Valid"
                  If (IsEmpty(Cells(j, 6).Value)) Then
            Cells(j, 6) = ActiveWorkbook.Worksheets("Zoom 2").Cells(i, 8).Value
                 End If
                Cells(j, 6) = ActiveWorkbook.Worksheets("Zoom 2").Cells(i, 8).Value
           Cells(j, 7) = Cells(j, 7) + ActiveWorkbook.Worksheets("Zoom 2").Cells(i, 5).Value
        
            End If
        
     Next j
 Next i
 
   For i = 2 To 200
           For j = 2 To 128
       
               If (ActiveWorkbook.Worksheets("Zoom 3").Cells(i, 1).Value = Cells(j, 1)) Then
         Cells(j, 4).Value = "Present"
         ActiveWorkbook.Worksheets("Zoom 3").Cells(i, 10).Value = " Valid"
                  If (IsEmpty(Cells(j, 6).Value)) Then
            Cells(j, 6) = ActiveWorkbook.Worksheets("Zoom 3").Cells(i, 8).Value
                 End If
                Cells(j, 6) = ActiveWorkbook.Worksheets("Zoom 3").Cells(i, 8).Value
           Cells(j, 7) = Cells(j, 7) + ActiveWorkbook.Worksheets("Zoom 3").Cells(i, 5).Value
        
            End If
         
     Next j
 Next i
 
 
   For i = 2 To 200
           For j = 2 To 128
       
               If (ActiveWorkbook.Worksheets("Zoom 4").Cells(i, 1).Value = Cells(j, 1)) Then
         Cells(j, 4).Value = "Present"
         ActiveWorkbook.Worksheets("Zoom 4").Cells(i, 10).Value = " Valid"
                  If (IsEmpty(Cells(j, 6).Value)) Then
            Cells(j, 6) = ActiveWorkbook.Worksheets("Zoom 4").Cells(i, 8).Value
                 End If
                Cells(j, 6) = ActiveWorkbook.Worksheets("Zoom 4").Cells(i, 8).Value
           Cells(j, 7) = Cells(j, 7) + ActiveWorkbook.Worksheets("Zoom 4").Cells(i, 5).Value
           
        
            End If
         
     Next j
 Next i
 
            For j = 2 To 250
             If (IsEmpty(Cells(j, 4).Value)) Then
                Cells(j, 4).Value = "Absent"
            
             End If
           Next j
      For j = 2 To 250
               If (Cells(j, 4).Value = "Absent") Then
                  Cells(j, 6).Value = "N/A"
                   Cells(j, 7).Value = "N/A"
               End If
          Next j


End Sub
