# stock-analysis #

## Project Overview

Analysis of 12 renewable energy stocks using VBA in Excel to see the total daily volume and return through 2017 and 2018.

## Results

* After creating a Macro that went through each stocks ticker to compile their total daily volume throughout the year and their return. We found that each stock seemed to preform better in 2017 than 2018, with the exception of $ENPH and $RUN
* 2017 Stock Analysis:
   ![Screen Shot 2022-05-30 at 12 05 25 PM](https://user-images.githubusercontent.com/96406929/171048688-f57eef69-b34f-4179-a562-4744f4d1f99c.png)
* 2018 Stock Analysis:
   ![Screen Shot 2022-05-30 at 12 05 37 PM](https://user-images.githubusercontent.com/96406929/171048717-be2347ac-d771-40e7-9fb2-a180a406b8ba.png)
* Another result we found that was when we refactored our analysis the run time of our refactored script was significantly faster than the orignial script, evidenced by the images below
  * Original 2017 Stock Analysis Run Time:
     ![VBA_2017_Original_Return_Time](https://user-images.githubusercontent.com/96406929/171049424-3d718122-1d62-4b90-b4ec-4616ad9cd8e4.png)

  * Original 2018 Stock Analysis Run Time:
     ![VBA_2018_Original_Return_Time](https://user-images.githubusercontent.com/96406929/171049450-9d2eb3e1-eaa3-42b7-8fa1-c648cfc391f5.png)

  * Refactored 2017 Stock Analysis Run Time:
     ![VBA_Challenge_2017](https://user-images.githubusercontent.com/96406929/171049463-6582d702-8693-4f1f-8da2-6b4ac9285a2d.png)

  * Refactored 2018 Stock Analysis Run Time:
     ![VBA_Challenge_2018](https://user-images.githubusercontent.com/96406929/171049474-8d96ffc2-b070-4b85-8080-26f2c1ea1d61.png)

## Summary 

Following the analysis it was discovered that each stock preformed well in 2017 with the exception of $TERP, while 2018 saw only $ENPH and $RUN having positive returns. I also found that some of the pros of refactoring code are making the code more applicable to different circumstances and optimizing run time, while some cons are debugging and finding the less optimal options made the first time around. In application to our VBA script these pros were seen in the return times greatly decreasing and making the code more flexible to include different years we could analyze. Cons I encountered were human errors made when it came to spelling and syntax that lead to instances of debugging. 
