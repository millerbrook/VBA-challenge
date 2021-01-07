# VBA-challenge
Homework #2 for UCF Data Bootcamp
I enjoyed completing this challenge and hope to work on it to add features such as a dashboard with user-selected charting.

Here are some peculiarities of how I completed the homework:
a. at over 400 lines, it seems long. See b-f.
b. one part of this challenge was a structural dissimilarity between the test .xls and the actual .xls. Late in the process I realized it would be easy to create a more useful test .xls, which helped with troubleshooting (as well as designing a script to reset the .xls after the main script had run!)
c. the specific challenge: in the test .xls any given stock only appears on one sheet. so,  the script creates a 'yearly totals' sheet, and once that was populated creates a 'three year totals' sheet. the latter does math and formatting only on the yearly totals sheet.
d. this contributed greatly to length. I essentially copied the main internal nested loop and repurposed it to make thethree year totals sheet. If preferable, the yearly totals sheet could subsequently be hidden or deleted through a couple of lines of code. Howver, having both creates opportunities for interesting analyses of the data.
e. another challenge: 'initial open' needed to be captured at the right moment in the loop, and I decided it was easiest to store it in the two new sheets (see 'a'). again, this opens the door to some additional analysis.
f. a final challenge: calculating % increase for stocks with 0 as an initial open value. I settled on '0' as a default value, which does NOT provide a sense of the growth of certain stocks. I'd want to have guidance about best practice for this.
g. finally, I also went bananas with comments. This helped me stay organized, but is probably not feasible in really long programs. I also used highly specific variable names. This helps me, but makes the code dense to read. 

Thanks for taking the time to look through my code. Advice about any of these issues, and about cool 'next' features to build if I want to make this part of a showcase repo, would be helpful.
