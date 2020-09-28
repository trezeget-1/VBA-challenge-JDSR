# VBA-challenge-JDSR
This is the 2nd hwk for the Data Analytics Bootcamp

To run the code, you have to open Visual Basic in Excel, place the cursor in the first macro and run it, it will automatically run the 2nd macro too.

I added a rule in my code in which if there was a Ticker in which the openning price was 0 (zero) then I would consider the opening price as 0.01.
The reason for this is that I need to divide the Yearly Change by this value and there was one ticker in the sheet 2015 in which the opening price was 0 and its yearly change value was 15, so it would have been a wrong statement to say that there was no yearly percentage change for this ticker, but calculating it with an opening price of 0 (zero) would have been imposible, so I made this assumption.

Thank you!
