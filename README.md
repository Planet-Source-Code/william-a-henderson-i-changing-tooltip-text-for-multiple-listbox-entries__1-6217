<div align="center">

## Changing tooltip text for multiple listbox entries


</div>

### Description

The purpose of this code is to display a tooltip giving a description of each entry in a listbox.
 
### More Info
 
Standard inputs relative to the MoveMouse event.

Familiarity with the ItemData property of the listbox control.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[William A\. Henderson I](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/william-a-henderson-i.md)
**Level**          |Advanced
**User Rating**    |3.0 (9 globes from 3 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[VB function enhancement](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/vb-function-enhancement__1-25.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/william-a-henderson-i-changing-tooltip-text-for-multiple-listbox-entries__1-6217/archive/master.zip)





### Source Code

```
'FACED WITH THE PROBLEM OF SHOWING OFFICE CODES
'FOR OFFICES WITH DUPLICATE NAMES IN A LISTBOX,
'AND NOT WANTING TO INCUDE THE NUMBER IN THE
'TEXT ENTRY IN THE LISTBOX, I DEVELOPED A QUICK
'WAY OF SHOWING THE NUMBER WHICH WAS STORED IN
'THE LISTBOX ITEMDATA PROPERTY.
'
'NOTE:
'WordHeight = 195 (depending on the font used).
'
'THIS CODE IS AN IMPROVEMENT UPON CODE PREVIOUSLY
'SUBMITTED BY ANOTHER VB PROGRAMMER INWHICH THE
'PROGRAMMER LOOPED THROUGH EVERY ITEM IN THE
'LISTBOX TO DETERMINE WHICH TEXT TO DISPLAY IN THE
'TOOLTIP. THE PROBLEM ENCOUNTERED BY THAT CODE WAS
'THAT IT DID NOT WORK FOR LARGE LISTBOXES WITH
'ENTRIES GREATER THAN 167. ON THE 168th ENTRY, AN
'OVERFLOW ERROR WAS ENCOUNTERED. MY CODE IS FASTER
'AND TAKES YOU DIRECTLY TO THE ENTRY WITHOUT
'LOOPING THROUGH THE LIST.
'
Private Sub ListBox1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim index As Integer
  index = ListBox1.TopIndex + ((Y) / WordHeight)
  ListBox1.ToolTipText = Str(ListBox1.ItemData(index))
End Sub
```

