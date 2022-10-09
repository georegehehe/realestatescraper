# realestatescraper
## Introduction
This projects delves into the recent alleged problem of the high default rate of real esatate projects in China. 
Specifically, two python programs have been developed using Selenium to collect numerical and textual data respectively regarding the issue for the past 
>commentscraping.py collects data since May from http://liuyan.people.com.cn/threads/list?fid=5063&position=1&utm_source=pocket_mylist, where comments from chinese
internet users regarding the issue can be found

>main.py is the more essential part of this project, which collects the Banker's Acceptance Credit data for all real estate companies with an overdue record during the past year.
Website can be found here: https://disclosure.shcpe.com.cn/#/infoQuery/ticketStateQuery

>违约房地产公司2021-2022.xlsx: input file that contains the companies with overdue records during the indicated years

>result.xlsx: output of main.py that summarizes overdue/acceptance amount for the above companies

>comments.xlsx: output of commentscraping.py that summarizes the contains the post date, the title, the id, and the main body of the comments

## Special Features
>commenscraping.py utilizes multithreading to speed up accessing hundreds of websites. 

>main.py implements an image recognition function called "find_offset", which is able to solves the slider captcha problem given by the website and therefore access 
the information. The original image is being pooled into a smaller matrix, where the average rgb value is being calculated for every 20 by 34 block (the size of the
missing pattern). Since all missing patterns are white, the lowest average rgb value is calculated for the matrix, and its y value is taken as the offset for which the
slider needs to be moved in order to solve the problem. Until now, the accuracy for this function is still 100% on this particular pattern

>defensive programming is applied due to the volatile nature of selenium

## Notes

Data for reference/educational purposes only. It does not imply any political views from the author

## References
https://stackoverflow.com/questions/57060448/cannot-take-screenshot-with-0-width
https://blog.csdn.net/Mingyueyixi/article/details/104345623

