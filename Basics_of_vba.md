# Create Macro

> To integrate vba code to excel you need to save them as Macro, these macro can then later be linked to any type of action, such as __button__ and keyboard shorcut.

## Change the value of the active cell with __Test__.
~~~
ActiveCell.Value = "Test"
~~~

## Change the __Fill__ color of the entire selection with __Red__.
~~~
Selection.Interior.Color = RGB(255, 0, 0)
~~~

## Change the value from cell __A1__ to _B3_ with _Test_
~~~
Range("A1:B3").Value = "Test"
~~~

## The Offset Property adjusts your position based on the initial Range you define.
~~~
Range("A1:B3").Offset(1, 2).Value = "Test"
~~~

## Change the value of cell __A4__ with the __Average__ from cell __A1__ to _B3_
~~~
Sub Macro_Name()
    Data = Range("A1:B3").Value
    Range("A4").Value = Application.WorksheetFunction.Average(Data)
End Sub
~~~

## Change the value from the __Min__ and __Max__ of a column with a random value between __Min__ and __Max__ then average the result in the column bellow
~~~
Sub Macro_Name()
    Dim i As Integer
    Dim Max As Integer
    Dim Min As Integer
    Dim random As Integer
    Randomize ' Initialize random-number generator.
    Max = 10
    Min = 1
    For i = 1 To Max
        random = Int((Max * Rnd) + Min)
        Cells(i, 1).Value = random
    Next i
    DATA = Range(Cells(Min, 1), Cells(Max, 1))
    Cells(Max, 1).Offset(1, 0).Value = Application.WorksheetFunction.Average(DATA)
End Sub
~~~

## Showoff your skills
~~~
Sub Macro_Name()
    Dim index_x As Integer
    Dim index_y As Integer
    Dim Max As Integer
    Dim Min As Integer
    
    Dim random_r As Integer
    Dim random_g As Integer
    Dim random_b As Integer
    
    Randomize ' Initialize random-number generator.
    Max = 38
    Min = 1
    For index_y = Min To Max
        For index_x = 1 To Max
            random_r = Int((255 * Rnd) + 0)
            random_g = Int((255 * Rnd) + 0)
            random_b = Int((255 * Rnd) + 0)
            Cells(index_y, index_x).Interior.Color = RGB(random_r, random_g, random_b)
            Cells(index_y, index_x).Value = "I ROCK"
            
            'random_r = Int((255 * Rnd) + 0)'
            'random_g = Int((255 * Rnd) + 0)'
            'random_b = Int((255 * Rnd) + 0)'
            'Cells(index_y, index_x).Font.Color = RGB(random_r, random_g, random_b)'
            'Cells(index_y, index_x).Font.Bold = True'
            Cells(index_y, index_x).Borders.LineStyle = xlContinuous

        Next index_x
    Next index_y
End Sub
~~~

Source: [select-and-selection](https://wellsr.com/vba/excel/select-and-selection/) 

You can find more documentation on documentation vba in excel on [wellsr](https://wellsr.com/vba/excel/).


# Optional
## How to activate the _dev_ pannel.
Check the dev option in the list located in:
~~~
File/Option/Customize_Ribbon/Customize_Ribbon
~~~
> The path name might be different if excel in other language then english.
