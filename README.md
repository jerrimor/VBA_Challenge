# VBA_Challenge
Refactoring code for a macro so that the results of the analysis will run efficiently regardless of data size.  

##Overview of the VBA Challenge Project
Steve has asked for analysis help on stock perfomances so that he can assist his parents in their investment making decisions. The parents had one stock in mind however Steve has access to more data on different stocks.  The analysis will be run on multiple stocks to compare and realize those that perform best. A macro was created to sort through the data and display the ticker symbol, Total Daily Volume and the Return on each ticker.  In this challenge specifically, additional code was added to the macro so that it would run more efficiently so that it can be used in the future for larger data sets. 

##Results

The stock analysis that can be deduced from this macro is a comparison between two years, 2017 and 2018, for the list of stock indexes provided.  As visualed below, most of the stocks had a negative return in 2017 (as seen as red cells) and a positive return in 2018 (as seen by the green highlighted cells).  The volume demonstrates the stocks that performed well in sales and the apparent winner is the SPWR, Sun Power Corporation in both years.  This would be an ideal stock to purchase so that the future sale would be quick and profitable.  

![E30F2F21-4E46-4DE5-A6E0-AAB4C2B711CA](https://user-images.githubusercontent.com/96222437/148702938-213eaedf-6cc6-4bdb-8896-42c6eb84df02.jpeg)
![8099823D-7D1A-4A38-8C49-9338289D1A62](https://user-images.githubusercontent.com/96222437/148702941-0127bcb6-e487-4088-be22-167ffb16298b.jpeg)


Refactoring the code in this challenge proved to help the run time of the macro.  Although the decrease in run time appears to be minimal and a small value, 2017 running initally in 0.2871094 seconds and then after the refactoring, 2017 runs in 0.2773438 seconds.  Screen prints are below to show the displayed run time amounts between the original code and the refactored code.  These minimal improvements will increase as the data increases.  

*Run times for Original Code* :smirk:

![9C6581BB-D327-4AB9-A70F-0B438838BFC2](https://user-images.githubusercontent.com/96222437/148701096-18dae7fa-f6f9-4614-92f6-a97d58cacb62.jpeg)

![D68B1537-F03A-4AF4-BBB3-C5FBECF39060](https://user-images.githubusercontent.com/96222437/148701105-2c55b832-f0c9-4f4d-9742-c438743da7b8.jpeg)

*Run times for Refactored Code* :partying_face: :fire:

![4B416B7B-93BF-4A17-96C3-A9D1EFA218FB](https://user-images.githubusercontent.com/96222437/148701115-3f36c2a9-4c30-4cef-a141-ae515af74fd4.jpeg)

![549CCF19-0B14-4DE3-8ED5-FA48864FC06D](https://user-images.githubusercontent.com/96222437/148701125-04c4156a-2cc8-4127-974b-23744eab733b.jpeg)

##Summary

###Indexticker is here to save the day! 
Refactoring the code was accomplished by adding the IndexTicker index to hold the value of the row of the spreadsheet. Snippet of the code with the new value is below.  The new index is utilzied throughout the code, as seen below, it is referenced for the starting and ending prices, along with the ticker volume which creates a streamlined approach for capturing and storing those values. 

tickerIndex created and utilized within the code

![BD80176B-9419-4D88-82F4-65431E7F5B22_1_201_a](https://user-images.githubusercontent.com/96222437/148702482-189c33f3-db16-4743-bcdb-88feda446e4b.jpeg)




###Bye-Bye Nested For Loop!

The nested For loop was removed, and single loops implemented for efficiencies. 

*Original Code with Nested For Loop*
The first For Loop begins with 'i' and the closing of the Loop is not visible because it occurs after the second For Loop which is the 'j' loop.  The j For loop is nested within the i For loop.  It is a mathematical theory that nested For Loops cause inefficiencies and that organizing the single loops would bring about an increase in speeds for running code. 

![EC7040B8-8244-441F-974E-AEAF3DE782F1](https://user-images.githubusercontent.com/96222437/148701897-03ff4248-7604-4fd6-b8b1-591fcac1684a.jpeg)


*Refactored Code*

![F47F22A9-2036-42E5-81DB-692AFC87E082](https://user-images.githubusercontent.com/96222437/148701731-9245514f-32d8-4932-b08a-e0cd216bda15.jpeg)
