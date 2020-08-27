# Cellular Automata and Conway's "Game of Life"

This is an excel implemetation of Conway's "Game of Life"
This program is all included in The_Game_of_Life.xlsm 
and can be downloaded on to any computer that will run Micropsoft Excel.

## Rules of the Game:

Conway's "Game of Life" is a simulation that runs on a few simple rules that are 
supposed to mimic the idea of either machines or organisms that replicate themselves. 
Each unit is either created or sustained with the help of other units.

The rules are as follows:

1. Each Cell is either dead or alive and the next generation is determined by the number
of neighbors the cell has in the current generation.
2. If the cell is dead and has exactly 3 neighbors that are alive, 
the cell will be alive in the next generation.
3. If the cell is alive and has exactly 2 or 3 neighbors, 
the cell will be alive in the next generation.
4. If neither rule 2 or 3 are true then the cell will either die or remain dead in the next generation.

### Dificulties and Solutions

The major dificulty I had with this project was trying to get the animation to wrap around the edges of the 
simulation field. I was able to implement it in the logic of the algorithm but because of how VBA for Excel
is set up, the array I was storing my data in had an index base of 1 instead of 0 like most programming 
languages use.

In order to fix this I had to first copy the data out of the excel cells into a base-1 array and then 
loop through that array in order to move the data to a base-0 array. I was then able to loop through 
the new array to apply the needed algorithmic logic and applied it back onto the base-1 array so that it 
could be pasted back into the excel cells. While this wasn't ideal it didn't change the runtime of the 
function because it only ever has to loop through the array 2 times and in Big O notiaton the constants are 
dropped making the run time of O(2N) still O(N).

### TO-DOs I Would Like to Try
- One thing I did try is changing the color of the cell based on how many neighbors it has, but because of how I
implemented that, using conditional formatting, it slowed down the simulation greatly. This has been saved under 
The_Game_of_Life_v2.xlsm. I would be interested in trying that again, but in a way that didn't slow it down quite 
so much.
- Another thing I would like to try would be to set up the simulation so that each cell on the field has a limited
number of resources and once they are all used up, that cell would be permantly dead and would ignore the rules that
allow a cell to come to life.
- A variation on the above that also sounds interesting, would be to continue to set it up with limited resources,
but allow the resources to either grow one unit every so many generations, and/or have the death of a cell to add a
small amount of resources back to that space on the field.
