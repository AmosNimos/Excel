# Create Macro

> To integrate vba code to excel you need to save them as Macro, these macro can then later be linked to any type of action, such as __button__ and keyboard shorcut.

## Change the value of the active (main_selected) cell with __Test__.
~~~
ActiveCell.Value = "Test"
~~~

## Change the __Fill__ color of the entire selection with __Red__.
~~~
Selection.Interior.Color = RGB(255, 0, 0)
~~~

## Change the value of the cell from __A1__ to _B3_ with _Test_
~~~
Range("A1:B3").Value = "Test"
~~~

# Optional
## How to activate the _dev_ pannel.
Check the dev option in the list located in:
~~~
File/Option/Customize_Ribbon/Customize_Ribbon
~~~
> The path name might be different if excel in other language then english.
