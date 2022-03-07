# Refactor-VBA
Module two challenge---Refactor VBA Code and Measure Performance 
# Refactor VBA Code and Measure Performance


## Overview of Project

The challenge required editing or refactoring the code we did for the Stocks Analysis project in Module 2. The challenge asked for generation of a solution code to loop through all the data one time in order to collect the same information we collected in the previous module. Then, to determine whether refactoring the code successfully made the VBA script run faster. Finally, presentation of a written analysis that explains these findings.  


### Purpose

The purpose of the the challenge was to attempt make the code more efficient—by taking fewer steps, using less memory, or improving the logic of the code to make it easier for future users to read. The challenge requires a different approach to the previous module, in which the code loops through each row of data and loops through the tickers before moving to the next row. The challenge then uses the timer function to determine if the previous module or this one was faster---to compare the performance of each code stucture.

## Analysis and Challenges

We were given a starter code, as a guide, to run the code through a loop one time---rather than multiple times. For example, in the first module, we ran the code through every row of the sheet for each ticker---12 tickers x3013 rows x 2 sheets (2018 and 2017). On the refactoring challenge, the code ran through each ticker at each row and all the code at each row before moving to the next row, and only went through the sheets 1 time, instead of reading each sheet of data 12 separate times. This made the code faster and more efficient than on the previous challenge. I was able to compare the speed of each code run-through using the timer function (screenshots below).

For analysis, here are screenshots of the original runtime for the Stocks Analysis project:

https://github.com/lllohr/Refactor-VBA/blob/main/Resources/Stocks_Analysis_2018.png

https://github.com/lllohr/Refactor-VBA/blob/main/Resources/Stocks_Anaysis_2017.png


### Challenges and Difficulties Encountered

The biggest challenge for this project was refreshing my memory of the previous module. As I was going through each task, I looked to my code for the previous challenge and tried to reuse whatever code would accomplish my task.  The variable name changes did provide another challenge---how was this different and why was my task different. Once I understood the task, I began to create my psudeocode, without realizing it was already in the VBA_Challenge.vbs file. This turned out to be a great confidence booster, as I found my psuedocode matched nicely with what was included in the starter code file.  

I ran into some difficulties remembering how to set the tickerStartingPrices and tickerEndingPrices using the tickerIndex. Through my first run through, I was stumped and could not remember how we set up the starting and ending prices, so I set the if statement to if the the volumes were zero. :

	  If tickerVolumes(tickerIndex) = 0 Then

		  tickerStartingPrices(tickerIndex) = Cells(i, 6)

	  End If

	  If tickerVolumes(tickerIndex) = 0 Then

		  tickerEndingPrices(tickerIndex) = Cells(i, 6)

	End If

Fortuitously, the code rendered the correct values and was faster the the Stocks Analysis project. After analyzing this however, I was unhappy with the possible problems that could ensue. I cleaned up the code by reusing the code from the previous project instead---changing the variables as appropriate.

Here is a run how fast the refactored code was with the above code lines: 

https://github.com/lllohr/Refactor-VBA/blob/main/Resources/VBA_Challenge_2018.png

https://github.com/lllohr/Refactor-VBA/blob/main/Resources/VBA_Challenge_2017.png

With the replacement with the following code, my results were slightly slower than when I set the volume to zero, however, I feel like it was the proper way to write the code without running into some inadvertent errors. The code did break my brain, going through the nested if statements.

		'If the the first ticker=value in first column keep going
            If Cells(i, 1).Value = tickers(tickerIndex) Then
                    
                    '3b) Check if the current row is the first row with the selected tickerIndex
                    If Cells(i - 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
                    
                        tickerStartingPrices(tickerIndex) = Cells(i, 6)
                        
                    End If
                    
                '3a) Increase volume for current ticker
                tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
                
                
                'Step 3c if-then if current row is last row then assign tickerEndingPrices variable
                If Cells(i + 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then

                    tickerEndingPrices(tickerIndex) = Cells(i, 6)
                
                End If
                
            End If

The original screenshots were included in the project to show the speed difference between the code. Here is the runtime after modifying the code:

https://github.com/lllohr/Refactor-VBA/blob/main/Resources/VBA_Challenge_2018Mod.png

https://github.com/lllohr/Refactor-VBA/blob/main/Resources/VBA_Challenge_2017Mod.png

Another challenge I had was that my dates in the 2017 and 2018 sheets had somehow been corrupted when they were copied. I hadn't noticed it until the end of the project, but it had some formatting issues. The solution was to replace the corrupted sheets with clean data and run the analysis.

## Results

- What are two conclusions you can draw about the challenge? Did refactoring work? What are advantages and disadvantages of refactoring code in general?

From what I can conclude, the code was faster and more efficient after refactoring. I believe the benefits of refactoring are that tasks can be combined from multiple lines of code into one of two lines of code. Additionally, like in this instance, it used less memory, I assume to run this through the second way. We were able to think about another way to do the same task quicker and more efficiently. The advantage of refactoring, is that the code is already written, we know it works! We can ask ourselves, "How can we perform this better?" 

I was able to create the same results as I did previously in the first project:

https://github.com/lllohr/Refactor-VBA/blob/main/Resources/Output_2017.png

https://github.com/lllohr/Refactor-VBA/blob/main/Resources/Output_2018.png

I think a disadvantage to refactoring might be that you can take good code that works and potentially break something that isn't broken. Every time we alter code, we have the potential for mistakes. In a real world example, there have been "fixes" sent as updates to applications on my phone. In many instances, I was not encountering any issues prior to the fix and have new issues emerge once the fix is applied.

- What can you conclude about the project? What are some advantages and disadvantages of the original and refactored VBA script?

I think the advantages and disadvantages of refactoring VBA are similar to what the advantages and disadvantages are in general. For example, when I completed my project, my buttons no longer worked and some of the code that was included in the original project were no longer there. I had to reprogram my buttons by creating new macros which were not contained within the vbs starter code file. While the script did work more efficiently, there will always be potential that the new refactored code may not be more efficient. Will we create new issues? 
