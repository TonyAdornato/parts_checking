# parts_checking
Short program that correlates Durr spare part numbers to Ford's part numbers

Background:
In the paint shop we installed new sealer robots from Durr. During the installation they gave us a list of all of the spare parts with descriptions and part numbers. We had some similar robots already installed in the shop, but these were a new revision with some new parts and some that carried over between generations. So we knew that there were going to be a lot of parts that had already been entered into the internal Ford parts system. 

We needed to go through the list and check to see which parts needed to be added and which already had a Ford part number. But the list of spare parts was thousands of entries long and I didn't want to sit around all week and check each part manually. So I decided to write a program that would do it for me.

Looking at the program you will see that it uses the selenium module to control the web browser. It controls the browser and navigates to the web page where we can look up parts in the database. This is because the people over at IT did not want to give me access to query the database directly. Access to that is kept limited to a few people and frankly getting added to the list was more hassle than it was worth. That meant that I would have to use Selenium. Since this is a program that will only ever be used a few times, that seemed like a fine compromise to me.

