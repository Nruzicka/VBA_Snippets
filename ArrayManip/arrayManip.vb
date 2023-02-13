Option Private Module

Public Function rangeToArray(target As Range) As Variant
    'converts 2 dim range to 1 dim array
    'exploits the for-each loop to capture non-continuous ranges
    Dim myArray()
    ReDim myArray(target.Count - 1)
    Dim dye As Variant
    
    Dim i As Integer
    i = 0
    For Each rngcell In target
        dye = rngcell.Value
        myArray(i) = dye
        i = i + 1
    Next rngcell
        
    rangeToArray = myArray
End Function

Public Sub quickSort(vArray As Variant, inLow As Long, inHi As Long)'Uses mid-pivot quicksort alogorithm. Only works for 1 dimensional arrays. Credit to allexperts.com
  Dim pivot   As Variant
  
  'pointer variables
  Dim tmpSwap As Variant
  Dim tmpLow  As Long
  Dim tmpHi   As Long

  tmpLow = inLow
  tmpHi = inHi

  pivot = vArray((inLow + inHi) \ 2)

  While (tmpLow <= tmpHi)
     While (vArray(tmpLow) < pivot And tmpLow < inHi)   'Dial in low pointer. Stop when it points to a value greater than pivot
        tmpLow = tmpLow + 1
     Wend

     While (pivot < vArray(tmpHi) And tmpHi > inLow)    'Dial in high pointer. Stop when it points to a value less than pivot
        tmpHi = tmpHi - 1
     Wend

     If (tmpLow <= tmpHi) Then           ' swap the values of the low and high pointers
        tmpSwap = vArray(tmpLow)
        vArray(tmpLow) = vArray(tmpHi)
        vArray(tmpHi) = tmpSwap
        tmpLow = tmpLow + 1                ' then move the pointers in.
        tmpHi = tmpHi - 1
     End If
  Wend                                     ' Repeat until the pointers collide into each other.

  If (inLow < tmpHi) Then quickSort vArray, inLow, tmpHi          'if the high pointer doesn't hit the lower boundry, then create a lower bound sub-array and repeat.
  If (tmpLow < inHi) Then quickSort vArray, tmpLow, inHi          'if the low pointer doesn't hit the upper boundry, then create a lower bound sub-array and repeat.
End Sub
