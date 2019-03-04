# Intermediate-Excel-WS
## The Essential Shortcuts
1. (Windows Only) You can use the alt key to manuever your Excel ribbon. 
Some key ones I use are:
  - HOI -> adjust column widths
  - HOA -> adjust row widths
  - AT -> creates filters for row headers
    - Hold Alt + down to access the filter bar without having to touch your mouse
  - AC -> Clear filter
  - HK -> sets cell format as number with commas
  - EL to delete sheets
    - SHIFT + F11 to make new sheets 
2. Moving around the Excel sheet.
  - Hold Ctrl + Direction keys in order to jump to an end of the data.
    - Holding shift along with this allows you to select multiple things quickly
  - F5 is the "GoTo" function which allows you to jump to specific cells.     
  - Home will bring you to cell A of whatever row you are on.
  - Page up and Page down allows you to scroll faster.
    - hold Alt while using page up and page down to scroll left and right faster.
3. etc
  - F2 allows you to enter edit mode for cells. 
  - F9 calculates the formula in the cell for you.
  - CRTL + Z to go back 
  - CRTL + Y to go forward
  - CRTL + X to cut
    - Be careful this one will cut all formatting as well
## The Pastas for copying
- Right click once you've copied something onto your clipboard. 
  - Bread and butter is paste values to avoid grabbing formulas that will give u gibberish.
  - Transpose is also useful
  - Link paste is useful when you want values to be combined to different sheets and update automatically
    - This leads to "Update Link" Messages when you open your worksheet.
ShortCut instead of right clicking is CRTL + ALT + V, but i actually prefer right clicking.
    - This option will also allow you to use the subtract/add paste which is useful for comparing numbers.
## Name Manager
The Name manager unde the "Formulas" ribbon allows users to manage specific cells by assigning them specific names. 
This is useful for managing other excel formulas(ex: Vlookup, offset).
You can quickly access all existing names in a worksheet by pressing the down arrow at the top left of your excel worksheet. 

## Vectorizing
### The CSE
Pressing Control + Shift + Enter allows you to enter formulas as arrays. This gives you the flexibility to remove intermediete steps when multiplying arrays. Some formulas, however, do require entering as an array which we will see later. 

## Functions
### Index 
The Index function allows you to index a specific array of data. Format: Index(Array,RowNum,ColNum)

### Match
Match allows you to find the the row/col number of a specific data point you might want from a list or array. Format: Match(item_youwant,list,type) Where type refers to whether you want the match to be exactly(=0) or "approximately"(=1). (I never use an approx match). 
I like to use this functions primarily as a list comparison checker. Problem: when you have multiple lists and you want to find which values list A might have that are also in list B. 

### Together 
Now as one might guess, these functions naturally work together and will end up giving you something similar to a vlookup or hlookup, but much more flexibile. The limitations of vlookup forces the data you are matching to have the indexed column on the left whereas, the Index-match combo has no restriction. Format: Index(Array_youWant,Match(index_var,Array_youWant,0))
You can now even add another match to index from a table rather than just an array!

### Offset
This functions allows you to look up the cell that may be i rows and j columns away from you with Offset(cell,i,j). I mainly use this function when iterating in vba since it's very useful for managing rows of data with one main cell as your anchor point. 

### Indirect
Indirect allows you to acces a cell by looking at another cell as a reference. This allows you to create control sheets where you may have a list of cells. 

### RAND
Gives you a random number in (0,1)
Format: Rand()

### LOGEST
For the regression analysis: y = b*m^x

Entering as a single cell will only return m whereas entering as an array will give you all the other regression statistics(which you can find by googling LOGEST, for this example I will only show the mechanics of comparing LOGEST and GROWTH so I only use the first row which is m and b respectively.) 
Generally I only use it as a single cell since all I care about is the trend.  
Format: LOGEST(Y-values,x-values,b_preference,stats_table)

### GROWTH
This function will give you the actual expected y's which are calculated. 








    
