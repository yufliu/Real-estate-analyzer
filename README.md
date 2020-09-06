# Real-estate-analyzer

Inspiration
I have been learning a lot about real estate and the ability to invest my future savings to set myself up better financially in the future. Real estate can be a great vehicle for wealth when done right but can cause major headaches and financial turmoil if not done carefully. I feel that it is very important for students to learn about investing their money into assets and vehicles of wealth and we do not learn enough about it in school leading to uneducated and potentially costly decisions in the future. One very important aspect of learning real estate is understanding the numbers and analyzing it before even considering what properties to buy. Analysis can take 5-20 minutes per property depending on your methods and knowledge. Analyzing 10 or 20 houses can really eat the time that people have to spend doing other things so I created this tool to help analyze houses much quicker and provide an easy and aesthetic platform to view the numbers behind a real estate deal. I hope this tool not only teaches but provides an easy and exciting way for people to get started in investing in real estate

What it does
This application takes basic user inputs such as down payment, interest rate (assuming a good credit score and a 30 year fixed rate), as well as incidents such as maintenance and vacancy rates. It then asks for a Zillow URL (any zillow url of a house for sale) and it will begin analyzing the property based on the zillow numbers. Once it is done getting the numbers from the listing, it highlights the values that will be used in yellow on the zillow webpage and outputs an excel sheet with all the numbers (rent, expenses) and the analysis of whether it is a good property or not (green or not and a check for pursue) including a pie chart of all the expenses that will go into the particular property. It then also outputs a website for the user to learn more about real estate terms which are also in the excel sheet in case they are not well versed in real estate.

How I built it
I used python to web scrape the information from Zillow using Selenium and then output it into an excel sheet using xlsx writer. I learned how to use plotly and how it can be used to output a pie chart. I learned how to input python strings and values into excel and format/modify existing cells to give it color and better text aesthetics. I had to re-learn html and css to help and modify the website template. I learned how to use various other libraries in python.

Challenges I ran into
Figuring out how to output the pie chart was difficult and took too much time. Figuring out how to highlight web elements surprisingly had me stuck for a bit and debugging all the errors I was getting because my xpath was missing variables or simply did not work.

Accomplishments that I'm proud of
I am proud that I was able to use selenium to quickly scrape zillow for the necessary information despite the 1000 times that it failed due to improper xpath.

What I learned
I learned how to use selenium to web scrape any website for information of use. I learned how to highlight web elements on the website. I learned how to modify existing templates in order to create my website.

What's next for Real Estate Analyzer
Next is to integrate the real estate analyzer to a website so any person can remotely access it. Will look into flask or django. Wanted to do it in this one but I simply did not have enough time. Also, want to improve the website to allow for a better interface and more information as I also ran out of time to customize it more. Will also be looking into integrating an API for a more seamless experience.
