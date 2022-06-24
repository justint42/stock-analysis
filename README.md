# Stock Analysis - Refactoring Code in VBA
## Background
A close friend named Steve has asked for assistance in creating an Excel macro to analyze the stock market for the years 2017 and 2018. Finishing this would allow his parents to make more infomred investment decisions in the future.
### Purpose 
The purpose of this project is to practice the user's ability to analyze a multitude of data with efficient code. While the current method works, there is better, more effective code to accomplish Steve's goal. Refactoring existing code allows it to take fewer steps, use less memory, and improve its logic so that future readers may interpet it more easily.

## Results

The lesson module required the user to create a nested for loop. In comparing the original code to the refactored template, I found that the efficiency difference relied on modifying the nesting order of my for loops. Creating the designated ticker arrays allowed the tickerIndex to access the relevant indexes in the code. Below is a comparison of main code differences. 

### Original Code

<img width="694" alt="Original Code" src="https://user-images.githubusercontent.com/106895220/175461734-7961aec5-00b3-47da-b775-ca49656c7e91.png">

Designating variables such as tickerIndex, setting up a solid relationship between the three ouput arrays and the iteration of tickerVolumes, looping over all the rows in the spreadsheet, and placing the output on its own outside loop reduced the number of iterations. 

### Refactored Code

<img width="692" alt="Refactored Code" src="https://user-images.githubusercontent.com/106895220/175461739-63e24739-84a2-4d75-acee-4e96134192ea.png">

### Original Completion Times
<img width="417" alt="VBA_Challenge_2017(1)" src="https://user-images.githubusercontent.com/106895220/175462865-f5cf50a5-5755-4c6a-a0b8-1812bcdd33af.png">

<img width="419" alt="VBA_Challenge_2018(1)" src="https://user-images.githubusercontent.com/106895220/175462886-486ff1d1-60bb-4756-bd9e-222e8af03605.png">

### Refactored Completion Times*

<img width="255" alt="VBA_Challenge_2017(2)" src="https://user-images.githubusercontent.com/106895220/175462906-6da97783-892b-4dba-8640-e9e3639c364c.png">

<img width="257" alt="VBA_Challenge_2018(2)" src="https://user-images.githubusercontent.com/106895220/175462916-da3ac734-f981-4d75-8f36-ae29969bf076.png">

*Look of refactored completion times are different due to upgrading iMac from macOS Mojave to Monterey.

## Summary 

### Refactoring Advantages and Disadvatnages

Refactoring code is analogous to “trimming the fat” or redesigning a machine to do its job better. Eliminating unnecessary loops and repetitive/redundant code can dramatically improve code structure without modifying its fundamental functions. Why make something unnecessarily complex? However, tight deadlines, new development planning, the cost of refactoring, and the chances of creating a mess may prevent one from refactoring.

### Refactored VBA Script Advantages and Disadvantages 

In retrospect, there was little reason to have as many loops as we did in the original code. One huge pro is the ability to use the original code as a resource for refactoring; one can see where there is excess and redundancy. A potential con is a lack of understanding of syntax (a challenge that I later overcame). If the user does not have a strong understanding of what the code is saying, then he will not know how to make it better (like spoken language, how can one condense their speech if they do not have a fundamental understanding of English?).  The refactored script may also be too long and cost even more memory if the programmer does not know what they are doing (I encountered situations where I basically broke even when comparing my execution times). 
