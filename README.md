# **Stock-Analysis**

## **Module 2 Challenge** 

## *Overview of Project*

### Refactoring

In this challenge, we refactored the Module 2 solution code to loop through all the data one time in order to collect the same information we collected while working through the module. We will determine whether refactoring the code successfully made the VBA script run faster. 

### Stock Analysis

In addition to improving our code, we analysed the performance of Steve's parent's stocks! We will take a look at how the stocks performed in 2017 vs. 2018.

## *Results*

### Refactoring Results

By modifying our code we were able to improve runtime.

#### 2017 Original Code Vs. Refactored Code

![Runtime_2017_Original](https://user-images.githubusercontent.com/111163264/188522670-e420fd8b-fd44-4ad5-b2eb-61078893f1f9.PNG)

Our original code ran for 0.625 seconds while analyzing the data for 2017.

![VBA_Challenge_2017](https://user-images.githubusercontent.com/111163264/188522742-aef6a793-ec10-4502-b158-1c8542d5987b.PNG)

Our refactored code ran for 0.046875 seconds while analyzing the same set of data. We redused the time the code took by 0.578125 seconds. That is a 93% improvement.

#### 2018 Original Code Vs. Refactored Code

![Runtime_2018_Original](https://user-images.githubusercontent.com/111163264/188522973-6a5d7c9b-f602-47f2-848c-e5d0fed410b3.PNG)

Our original code ran for 0.6289063 seconds while analyzing the data for 2018.

![VBA_Challenge_2018](https://user-images.githubusercontent.com/111163264/188523030-27add1bc-9094-447f-83af-8f5b13be6871.PNG)

Our refactored code ran for 0.046875 seconds while analyzing the same set of data. We redused the time the code took by 0.582188 seconds. Again, we achieved a 93% increase in speed.

### Performance (2017 Vs. 2018)

This exercise is all about learning to use VBA; however, analysis is the name of the game. So lets take a look at how these stocks performed.

#### 2017 Stock Performance 

![2017_Analysis](https://user-images.githubusercontent.com/111163264/188523592-060d039a-045d-4f36-9919-65b788d02d42.PNG)

In 2017 we saw positive returns on all the stocks except on, TERP. 

![2018_Analysis](https://user-images.githubusercontent.com/111163264/188523694-f280b95f-ca76-45de-81ab-d2cfa7c32c11.PNG)

2018 has a different story to tell. Only two of the stacks, ENPH and RUN, saw a positive return. We would need to do some additional research to determine why we such a dramatic shift. 

## *Summary*

Refactoring is a necessary part of coding. As we learn and improve in our craft it is necessary and beneficial to revisit code to see if it can be improved or optimized. However, in coding as in life, sometimes we get it right the first time around. When we refactor code we run the risk of damaging or adding unnecessary complexity. Wisdom... We also have to take into consideration the amount of time it will take to refactor vs the benifit we hope to achieve.

As far as our code is concered, we were successful at decreasing the amount of time it took to run. However, I am new to coding and refactoring this bit of code took me a couple of hours. I will need to run that code 14,400 times just to break even. Advantage, I learned and practiced the skills taught in the module. Disadvantage, it wasn't necessary at all.  
