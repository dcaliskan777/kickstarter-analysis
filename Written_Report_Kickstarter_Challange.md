 # Kickstarting with Excel

## Overview of Project

Louise’s play Fever came close to its fundraising goal in a short amount of time. Now, she wants to know how different campaigns fared in relation to their launch dates and their funding goals. In this project, using Kickstarter data set which is an excel document with a large amount of cells, the campaign outcomes based on their launch dates and their funding goals are being analyzed and visualized.

Firstly, the relation between the success of campaigns in the parent category theater and the launch month is investigated and visualized by a line graph. Secondly, the relation between the success of campaigns in the subcategory plays (which is subcategory of theater) and their funding goals is analyzed and visualized by an other line graph.

In this report, Analysis of Outcomes Based on Launch Date and Analysis of Outcomes Based on Goals are discussed; challenges and difficulties encountered during the study together with how they were overcome are explained.

Finally, conclusions driven from the graphs, the limitation of the dataset and what can be done more are discussed.

### Purpose

The purpose of this project is to answer the following two questions:

First, what is the best month(s) to launch a theater campaign in, according to the dataset given in the excel document?

Second, what is the best goal interval(s) to arrange a theater/plays campaign, according to the dataset given in the excel document?


## Analysis and Challenges

### Analysis of Outcomes Based on Launch Date

In order to analyze Outcomes based on the launch date a pivot table and a pivot chart based on that table were used. The original data did not have parent category column so it was created by using category and subcategory column first. Because it was needed to to be filtered according to theater in the pivot table. Launch date was given originally in the unix timestamps format so it was not readable, it is converted to day/month/year format by using the formula 

> =(((J2/60)/60)/24)+DATE(1970,1,1)

and then the format was changed from General to date. Here, J is the Column headed as "Launched_at" and the created new column was headed as "Date Created Conversion". From this column two more column were created: first "Years" column was created by using the formula

> = YEAR(Q2)

and second "Months" column was created by the formula

> = TEXT(Q2,"mmm")

where Q is the column headed as "Date Created Conversion". After all, concerning pivot table and pivot chart were created which are given in the sheet named "Theater Outcomes by Launch Date" in the excel document named ![Kickstarter_Challenge](https://github.com/dcaliskan777/kickstarter-analysis/blob/main/%20Kickstarter_Challenge.xlsx). The main finding is the following line graph:

![](./resources/Theater_Outcomes_vs_Launch.png)

In this graph it can be clearly observed that in every month the number of successful campaigns of theater is larger than that of failed campaigns. The successes in the May is the highest. There is high successes in June and July also. Although, the success in February and April are the same, the fail in February is lower than the fail in April. So, among February and April, February is preferable to launch a theater campaign. Similarly, the success in April and August are almost the same, the fail in April is lower than the fail in August, which means April is preferable to launch a theater campaign rather than August.This information with the percentage of success is  visualized in the following table:

|Outcomes        |May    |June   |July   |February |April  |August |
|:-----          |:-----:|:-----:|:-----:|:-----:  |:-----:|:----- |
|Successful      |111    |100    |87     |71       |71     |72     |
|Failed          |52     |49     |50     |39       |40     |47     |
|Canceled        |3      |4      |1      |3        |2      |4      |
|% of Successful |67%    |65%    |63%    |63%      |63%    |59%    |
 
In summary, It can be recommended to Louise the launch the campaign in May, if it is not possible it may be launched in June, July, February, April or August in the given order.

### Analysis of Outcomes Based on Goals

In order to analyze outcomes based on the goals, the range of goals was divided into subintervals as "less than 1000","1000 to 4999","5000 to 9999","10000 to 14999", "15000 to 19999","20000 to 24999","25000 to 29999", "30000 to 34999", "35000 to 39999", "40000 to 44999", "45000 to 49999" and "50000 or more".The original data did not have subcategory column so it was created by using category and subcategory column first. Because it was needed to put the condition "plays" (which is a subcategory of theater) in the **COUNTIFS()** function. Four conditions are argued in **COUNTIFS()** function:lower bound of interval in the column headed as "Goals", upper bound of interval in the column headed as "Goals", "successful" or "failed" or "canceled" in the column headed as "Outcomes" and "plays" in the column headed as "Subcategory". So that, the table in the sheet named "Outcomes Based on Goals" in the excel document named ![Kickstarter_Challenge](https://github.com/dcaliskan777/kickstarter-analysis/blob/main/%20Kickstarter_Challenge.xlsx) was created. Percentages of "successful", "failed" and "canceled" campaigns of "plays" are computed and displayed in the same table also. Based on the table the line graph named "Outcomes Based on Goals" was created. The concerned graph is the following:

![](./resources/Outcomes_vs_Goals.png)

One quick look in this graph shows that the success and fail in plays campaigns are equivalent if the amount of goal is between $ 15000 and $ 19999. If the amount of goal is less than 15000 there is more chance for the campaign being successful than being failed. And amazingly, the probability that a plays campaign being successful is larger than the probability of a plays campaign being failed when the amount of the goal is between $ 35000 and $ 45000. Let's put the trend of percentage of successful under the scope by the following table:

|The intervals of Goals  |The percentage of Successful |
|:----------------------:|:---------------------------:|
|Less than 1000          |76%                          |
|1000 to 4999            |73%                          |
|5000 to 9999            |55%                          |
|10000 to 14999          |54%                          |
|15000 to 19999          |50%                          |
|20000 to 24999          |45%                          |
|25000 to 29999          |20%                          |
|30000 to 34999          |27%                          |
|35000 to 39999          |67%                          |
|40000 to 44999          |67%                          |
|45000 to 49999          |0%                           |
|50000 or more           |13%                          |

In this table percentage of Successful shows decreasing trend, however in the intervals 35000 to 39999 and 40000 to 44999, it increases unexpectedly. This might be due to the outliers of goals. In fact, the goals has 28 outliers in the plays campaign, which means all goals in last three interval and 2 goals in fourth interval from last are outliers of outcomes. Which factors causes this result we do not now. There might be some goals written mistakenly, it might be because of some factors which were given in the dataset but were not analyzed in this project, such as backers_count; or it might be because of some other factor which were not considered in the dataset. These should be answered, but this is out of scope of the project.

According to this statistical information, it might be recommended to Louise to arrange a campaign such that the amount of goal is no more than $ 15000.


### Challenges and Difficulties Encountered

I started to make the project before completing the internet document for the first module. So I did not point out that the conversion of unix timestamps format to readable format was in the document. This was my first challenge to convert the launched_at column to readable date format. I googled to find out how I can do it, and followed the instruction that I found in the google, and solved the problem. It taught me that I should finish the document in the module before solving the challenge of the module.

Second challenge was encountered when I decided to create a column concerning months of the launch of campaigns, I googled again and find out how to take out the moths from the date. 

The third challenge was about using the **COUNTIFS()**function, arguments in the **COUNTIFS** are really confusable to understand. I tried several times to overcome this problem. An other problem with this function, it was time-wasting to write conditions concerning boundaries of intervals, there were a lot of cells do do it, I asked myself how I can do it by using shifting the  a function through a column down; and then found out the solution which is to write boundaries in different columns. One may see this in the column I and the column J in the sheet named "Outcomes Based on Goals" in the excel document named ![Kickstarter_Challenge](https://github.com/dcaliskan777/kickstarter-analysis/blob/main/%20Kickstarter_Challenge.xlsx). 

The last challenge was to write the report, it was first time that I was using markdown document. I solved this problem by patience, It took me a lot.

To be honest, It would be too difficult to decide in which ways and how I could do the analysis if it would not be explained the project in a an inquiry based, step by step manner. It was great help me. Thanks to whom prepared the instruction of the challenge!
 

## Results

By the analysis of Outcomes bases on Launch Date, it can be concluded that for a campaign of theater the best month of the year is May. In every month of the year the number of successful campaigns is more than the number of failed campaigns. In addition to this, June, July, February, April and August are appropriate months to decide to launch a campaign in the given order to be more successful.

If we consider the analysis of Outcomes based on Goals, an evidential conclusion can be made as the best amount of goals for more success chance is the one less than or equal to $ 15000. Although, the percentage successful for the amount of goals, from 35000 to 45000, is high enough; this does not fit with the general trend. We can not trust this information, it is needed to be more deeply analyzed in different perspectives.

The data set does not include statistics about how social media was used in campaigns and how many people were reached out during the campaign. These are two important parameters to analyze success of campaigns. They may assumed to be some limitations of the dataset.

Outcomes can be analyzed based on the duration of the campaign. It can be done by creating a new column by using deadline and the launch date. The number in cell of the new column is the difference of numbers in the corresponding cells of columns "deadline" and "launched_at". The created column is the column of duration. For this, a table and a line graph can be arranged as we did for analysis of Outcomes based on Goals. In the studies which we did, instead of line graphs we may use bar charts. We can analyze Outcomes based on Backers_count similar to the analysis of Outcomes based on goals that we did.

Additionally, for being more accurate we can make outlier analysis, the dataset is large enough so that some data may be included mistakenly, there might be outliers for some other reasons to be investigated. For example, in goals, lower quartile is 2000, upper quartile is 15000, inter quartile range is 13000 and so the lower boundary of upper outliers is 34500. In this column 569 elements are outlier. It makes 14 % of the entire column, this is enough large to be taken into account, am I wrong? Among these 569 outliers 28 of them are in the theater/plays campaigns. This might explain why the high value of percentage of successful in the goal between 35000 to 45000 does not follow the general trend. In fact they are all outliers, whatever reasons are.


**Dursun Caliskan**

*2/22/2022*
