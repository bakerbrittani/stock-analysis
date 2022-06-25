# stock-analysis
###Utilizing Visual Basic for Applications to Analyze Stocks

## Overview/Purpose

  The intention of this analysis was to determine which stocks performed well in 2017 and 2018. In doing the comparison all at once, with one looped code that would run through the entire data set we returned results that could easily be interpreted. 

## Results

  In order to quickly visualize the data in a legible format, refactoring included the clean-up and utilization of the For Loop as well as If-Then statements as shown in the image below. 

![IMAGE](/ForLoopIfThen.png)
	
  The formatting of the outcomes with a delineated header row allows the reader to easily determine the header titles and what each value means. By using a cell fill color in the background of the return percentages, the reader can quickly scan for comparisons, or overall yearly analysis. The format-specific code used is shown in the image below. 

![IMAGE](/Formatting.png)
 	
  If we begin our analysis by focusing solely on 2017, we can see just using our formatted colors that just one stock, TERP, performed unfavorably within the year. Overall the three stocks with the greatest return percentages for 2017 were DQ at 199.4%, SEDG at 184.5%, and ENPH at 129.5%. Moving to focus on 2018, there was a drastic difference in overall stock performance that year, with only two stocks remaining favorable, RUN at 84.0% and ENPH at 81.9%. ENPH fell by almost 50%, whereas RUN actually increased it's return percentage from 5.5% in 2017 to 84.0% in 2018. If I were to have to recommend a 'safe' stock to the client Steve and his parents, based on this analysis I would recommend RUN. 

![IMAGE](/2017-2018Comparison.png)
	
  
  The original code took approximately 0.5 seconds to run for either 2017 or 2018. After refactoring, the code ran in 0.12 and 0.13 seconds for 2017 and 2018 respectively. The run time is dramatically improved by the refactoring! 

![IMAGE](/Resources/VBA_Challenge_2017.png)

![IMAGE](/Resources/VBA_Challenge_2018.png)

## Summary
### Advantages/Disadvantages of Refactoring Code
  Advantages include overall speed enhancement as shown by our results above, using less memory, legibility enhancement for yourself and other possible future users, as well as making the code easier to maintain in future.
	Disadvantages include introducing more room for mistakes when making changes, or creating new errors causing retesting that can be time-consuming.
### Applied to refactoring the original VBA Script
  While refactoring I ran into new scenarios that required me to review the module and class activities to refresh my memory on how to write that particular function. I knew the outcome I wanted but thinking of a different way to obtain it when there was already a working code right in front of me was difficult. I felt it was almost worthwhile to start from scratch on certain parts. The obvious improvement in the run time was the measurable factor that was proven. If we needed to expand or add to this code in the future I would much rather work with the cleaner, refactored code as well. 
