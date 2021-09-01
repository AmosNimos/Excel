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
Sub test()
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

Source: [select-and-selection](https://wellsr.com/vba/excel/select-and-selection/) 

You can find more documentation on documentation vba in excel on [wellsr](https://wellsr.com/vba/excel/).


# Optional
## How to activate the _dev_ pannel.
Check the dev option in the list located in:
~~~
File/Option/Customize_Ribbon/Customize_Ribbon
~~~
> The path name might be different if excel in other language then english.
