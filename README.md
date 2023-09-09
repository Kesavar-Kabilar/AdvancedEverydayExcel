# Everyday Excel (Advanced) by Kesavar Kabilar

This portfolio encompasses the comprehensive body of work successfully accomplished by me (Kesavar Kabilar) in pursuit of the prestigious Everyday Excel (Advanced) Certification, conferred by the esteemed University of Colorado Boulder. Through rigorous academic dedication and the application of advanced Excel techniques, this compilation stands as a testament to the mastery achieved in the realm of Excel proficiency.

## Certificate Details

Coursera Certificate Details: https://www.coursera.org/account/accomplishments/verify/PEHFFAWU3YL7

## Summary of Skills

1. **Dynamic Formulas** : Utilized dynamic formulas to calculate Raw Score, Total Possible Score, Percentage and the "Score with the lowest 5 dropped" based on changing data.
2. **Data Validation (Dropdowns)** : Implemented data validation for selecting students' names, ensuring accuracy and ease of use.
3. **Conditional Data Validation** : Employed conditional data validation to restrict class and question selections in the Question Lookup area, enhancing user experience.
4. **Data Cleansing** : Removed unwanted elements like blank spaces, strings, and the value 9999 from the column, ensuring that it contained only numeric values.
5. **VBA Scripting** : Used VBA code to reset conditional drop-down lists, a crucial feature for usability.
6. **Data Integration** : Ensured seamless data integration between the "Dashboard" and "Recipes" tabs for accurate calculations.
7. **Dynamic Quantity Adjustment** : Ensured that the solution adjusted ingredient quantities automatically when new recipes were added.
8. **Dynamic Calculations** : Created formulas that are updated in real-time based on user inputs, allowing for a dynamic amortization schedule.
9. **Lookup Functions** : Employed lookup functions such as Index/Match, Xlookup, and Vlookup to determine units of measure for each ingredient based on user selections.
10. **Real-Time Data Transfer** : Used pointer formulas to transfer calculations from the "Inventory" sheet to the "Dashboard" sheet, providing up-to-date inventory information.
11. **Unit Selection and Calculation** : Set up data validation for unit selection and used the Haversine formula to calculate distances in miles or kilometres.
12. **What-If Analysis Tool** : Utilized Excel's "What-if Analysis" tool for one-at-a-time (OAT) sensitivity analysis.
13. **Dynamic Text Display** : Ensured dynamic text display in cell B11 to guide users during city selection.
14. **Handling Conversion** : Ensured that months were correctly converted to years for payoff time calculations.
15. **SEQUENCE Function** : Generated ranges utilizing the SEQUENCE function to spill and automatically update the necessary cells.
16. **Factor Influence Analysis** : Examined how variations in factors like cost, royalties, capital, sales, and interest rates impacted project valuation.
17. **Scenario Testing** : Systematically modified one factor at a time while keeping others constant to understand individual factor sensitivities.
18. **FILTER Function with NOT, ISBLANK, and ISNUMBER Functions**: Applied the FILTER function to isolate the data in the targeted column. I created a filtering condition to exclude blank cells and non-numeric values from the data.
19. **Linear Regression, **Standard Error and R-squared**** : Utilized matrix functions in Excel to calculate model coefficients for polynomial regression. Calculated standard error and adjusted R-squared values for regression analysis.
20. **Macro-Enabled Workbook** : Enabled macros in Excel to work with a provided dataset and create a functional dashboard.

## Top Hat Consolidator

The assigned task involved creating a dynamic Excel dashboard to analyze student performance data imported from the educational tool Top Hat. This project required several key functionalities: setting up data validation for selecting students' names, calculating Raw Score, Total Possible Score, and Percentage for individual student performance, computing the "Score with the lowest 5 dropped," dynamically calculating class averages for each student, implementing conditional data validation for class and question selection in the Question Lookup area, and ensuring all calculations and drop-down lists would update seamlessly with changes in the underlying data. This task presented a complex challenge in Excel data manipulation and required a thorough understanding of data validation, formula-based calculations, and VBA scripting for creating a comprehensive and dynamic educational assessment tool.

To solve this challenging Excel project, I followed these steps:

First, I set up data validation in cell C5 of the "Dashboard" tab to allow me to select a student's name from the drop-down list. This list was dynamically populated from the student names in column A of the "Data" tab, ensuring that if any students were added or removed, the drop-down menu would update automatically.

Next, I calculated the student's Raw Score, Total Possible Score, and Percentage in the "Overall Performance" area of the Dashboard. The Raw Score was the sum of all points across all questions in the "Data" tab. Total Possible Score was calculated as the total number of questions (columns in the "Data" tab) multiplied by 5 (5 points per question). The Percentage was then calculated as the Raw Score divided by the Total Possible Score. I ensured that these fields would dynamically update if new columns of data (new questions) were added or if columns were deleted.

To calculate the "Score with the lowest 5 dropped," I summed the five lowest scores (all out of 5), subtracted this sum from the Raw Score, and divided this result by the Total Possible Score minus 25. This calculation provided the desired result.

For the class averages, starting in row 21, column C, I calculated the average score for the student in cell C5 for each class. This involved finding all the scores for the students in that class and averaging them. Similarly, in column D, I calculated the entire class (all students) average for each class. Importantly, I ensured that these averages would dynamically update if new students were added or new questions were added to the "Data" tab. Only the classes with questions associated with them would be displayed.

In the "Question Lookup" area of the Dashboard, I implemented data validation in cell G5 to allow the user to select a class from the drop-down list. This list was dynamically populated based on the classes available in the "Data" tab. In cell G6, I used conditional data validation to restrict the user's selection of questions to only those assigned in the selected class. I also implemented a reset mechanism for the conditional drop-down list in cell G6 when the class selection in cell G5 was changed. This involved using VBA code and making the final project file a macro-enabled .xlsm file.

Finally, in cell C8 of the "Question Lookup" area, I calculated and displayed the class average for the selected question. This calculation dynamically updated as new data was added or if the user changed the class or question selection.

## Francesca's French Bakery

The assigned task involved creating an Excel-based solution to streamline ingredient ordering for Francesca's French Bakery. This project required careful data integration between two worksheets, "Dashboard" and "Recipes." It mandated dynamic formulas to calculate units of measure for various ingredients, based on user selections. Additionally, the solution needed to accurately compute the weekly ingredient requirements and account for the constraints of ingredient sizes available in the market. Furthermore, it had to facilitate the addition of new recipes while automatically adjusting the necessary ingredient quantities. Overall, the task aimed to optimize the bakery's ingredient procurement process, ensuring efficient and precise ordering, making it a comprehensive and dynamic Excel project.

To address Francesca's bakery ingredient ordering problem, I utilized Excel's various features and functions to create an efficient solution. Firstly, I focused on the "Dashboard" tab, where I allowed the user to input the quantities of batches for different baked goods. This data was linked to the "Recipes" tab, specifically to the "Batches needed" cells, which were essential for calculating the ingredients required.

To calculate the units of measure for each ingredient, I utilized lookup functions based on the selected ingredient from the drop-down lists. For example, if "Butter" was selected, Excel would display "cup(s)" in the relevant cells. These formulas needed to be dynamic and auto-updating, ensuring that any changes to the ingredients in the recipes would be reflected correctly. I also paid attention to cell L27, which updated the units of measure when a user selected different ingredients in cell J27.

The "Weekly Amounts Needed" for each ingredient (column F) were calculated based on the quantities of recipes in cells C6:C11. Similarly, the quantities that needed to be purchased were determined by subtracting the current inventory from the weekly requirements, taking into account that some ingredients could only be purchased in specific sizes (e.g., eggs in dozens, powdered sugar in 2-pound bags).

Furthermore, I ensured that the solution accommodated the addition of new recipes on the "Recipes" tab. The user could specify the quantity of the new recipe in cell C11 of the "Dashboard" tab, and this would automatically update the "Weekly Amounts Needed" and "Need to be Purchased" quantities.

In summary, my Excel solution seamlessly integrated data between the "Dashboard" and "Recipes" tabs, automatically calculated units of measure, adjusted ingredient quantities, and allowed for the addition of new recipes. This approach ensured efficient ingredient ordering for Francesca's bakery while accounting for the constraints of ingredient sizes.

## Historical Weather Lookup (Part B)

To solve the Historical Weather Lookup (Part B) project, I began by building upon the work I had done in Part A, where I successfully created a lookup for historical weather data. This lookup allowed the user to input a specific date, and the corresponding High Temperature, Low Temperature, Precipitation, and Snow data were displayed.

In Part B, I extended my analysis to find the record values for these weather parameters across all available years (1897 to 2020) for a given month and day. To do this, I utilized Excel's functions and data manipulation techniques.

For each weather parameter (High Temp, Low Temp, Precipitation, and Snow), I found the record value by identifying the maximum or minimum value across all years for the selected month and day. These record values were displayed in cell C16, and the year in which these extreme values occurred was shown in cell E18.

Additionally, I ensured that cell D16 displayed the appropriate unit of measurement based on the selected weather parameter. It displayed "degrees F" if High Temp or Low Temp was selected, and "inches" if Precipitation or Snow was chosen.

To provide a user-friendly experience, I implemented data validation in cell C14, allowing the user to select the weather parameter they wanted to analyze from a drop-down list. This list contained the exact spellings as specified in the project instructions, as the grader file tested these exact strings.

In summary, I built upon the foundation laid in Part A by extending the analysis to find record weather data for a specific month and day across multiple years. I achieved this by using Excel functions to find maximum and minimum values and displayed the results in the designated cells, considering units of measurement and user-friendly dropdown menus for parameter selection. This project presented a challenging but rewarding opportunity to apply advanced Excel techniques to real-world data analysis.

## Amortization Schedule with Extra Payments

The given task involves creating an Excel dashboard to manage an amortization schedule with extra payments for a loan or mortgage. The user inputs essential information, including the loan period, principal amount, and APR rate, which calculates the monthly payment. A drop-down menu allows the selection of extra payment frequency, with options ranging from monthly to biennially, along with an extra payment amount. The dashboard then computes and displays the total interest paid and payoff time for both scenarios: with and without extra payments. The payoff time is displayed in years and months, with months correctly converted to years if applicable. The savings from making extra payments are also calculated. The key challenge is to ensure all calculations are dynamic and instantly updated as input values change, providing a comprehensive tool for managing loan scenarios with extra payments.

To solve the "Amortization Schedule with Extra Payments" project, I began by setting up an Excel worksheet. In cells C3, C4, and C5, I allowed the user to input the Loan Period, Principle (total loan/mortgage amount), and Rate (APR), respectively. Using these inputs, I calculated the Monthly PMT (monthly payment) in cell C7.

Next, I created a drop-down list in cell G3 using data validation, which allowed the user to select the frequency of extra payments. Options included "(no extra payment)", "Month", "Two Months", "Three Months", "Four Months", "Six Months", "Year", and "Two Years". The user could also specify the extra payment amount in cell G4 (Extra payment AMT), and I assumed this payment frequency applied for the entire loan period.

In the "Amortization Schedule" worksheet, I performed intermediate calculations. This involved setting up an amortization schedule, incorporating extra payments, and calculating the Total interest paid and Payoff time for both scenarios: with NO EXTRA PAYMENTS and with EXTRA PAYMENTS. For Payoff time, I ensured that if the months equalled 12, it was converted to 1 year and 0 months, and added to the value for years.

Finally, in cell G11, I calculated the Savings by making extra payments, which was simply the difference between the total interest paid for the two options (the difference between cells G7 and H7).

I made sure that all calculations were done "live" so that any changes to the inputs on the Dashboard instantly and correctly updated the output calculations.

Once I completed these steps and ensured that all values were being calculated properly, I clicked on the GRADE button to check if my calculations were correct. The grader validated different scenarios to ensure accuracy, and if correct, I received the completion code for the project to input into the "Amortization Schedule with Extra Payments Submission Quiz" for credit.

## Real-Time Regression

The given task involves creating a real-time polynomial regression tool in Excel. Unlike traditional regression models, where users need to manually set up regressor variables and rerun regression analysis for different scenarios, this project aims to simplify the process. Users can copy and paste their x-y data into the spreadsheet, select the desired polynomial order, and instantly obtain model parameters. The challenge lies in making this tool adaptable to varying data sizes and polynomial orders while also correctly handling unused model coefficients. This project requires a deep understanding of the matrix approach to linear regression and dynamic array formulas in Excel to ensure that the model coefficients update automatically and accurately in response to user inputs.

The core concept of this project was to enable real-time polynomial regression for user-provided data. To achieve this, I implemented the matrix approach to linear regression using Excel's array (matrix) functions. This approach allowed me to calculate model coefficients (Beta values) in cells E5:E9 on the "Dashboard" tab automatically.

The key challenge was making the solution dynamic to adapt to varying data sizes and polynomial orders. I utilized dynamic arrays and followed the guidance from the screencast "Dynamic array hints for Real-Time Regression." This ensured that the model coefficients updated correctly even when the user changed the order of the polynomial or input data of different sizes into cells A2:B21.

Furthermore, I addressed a specific requirement regarding unused model coefficients. If the user selected an order lower than 4, I used the IF function, along with ISBLANK and NOT, to display "N/A" in cells E8 and E9 (Beta3 and Beta4). For instance, if the user chose an order of 2, these cells would correctly display "N/A."

Finally, I calculated standard error and adjusted R-squared, following the guidance from the screencast "How to calculate standard error and adjusted R-squared," ensuring a comprehensive solution.

In summary, I created a dynamic, real-time polynomial regression model in Excel, implementing the matrix approach and using various Excel functions, such as IF, ISBLANK, NOT, and dynamic array formulas. This solution successfully met the project's requirements and provided users with an efficient tool for polynomial regression analysis in Excel.

## Rental Car Inventory

The assigned task was to create an Excel workbook that serves as a rental center inventory management system. The goal was to allow users to add or remove items from a central database and view key inventory properties, such as rental costs and quantities, on a user-friendly dashboard. The primary challenge was to implement dynamic data validation dropdown lists, where the options available in one dropdown depended on the selection made in the previous one. Additionally, the solution needed to update in real-time when new items were added to the inventory. The task also required resetting the dropdown lists when a different category was selected and ensuring specific text was displayed in certain cells when data validation was reset. Ultimately, the project aimed to provide an efficient and interactive tool for managing rental inventory in Excel.

I utilized Excel's data validation feature to create dropdown lists for Category, Description, Brand, and Size/Model, making these menus dependent on each other for efficient selection. For instance, when I selected a Category, it dynamically updated the available choices in the Description dropdown, and similarly for Brand and Size/Model based on the prior selections.

To ensure the data updated in real-time, I used pointer formulas to transfer calculations from the "Inventory" sheet to the "Dashboard" sheet. This allowed me to display essential information such as the Cost per hour, Cost per day, Cost per week, and Quantity based on the user's selections.

I incorporated Visual Basic for Applications (VBA) to reset the conditional data validation drop-down lists whenever the user changed their selection, as described in the project requirements. If a different Category was chosen, the dependent dropdowns (Description, Brand, Size/Model) reset accordingly. Additionally, cells C3, C4, and C5 displayed "-- Choose Description --," "-- Choose Brand --," and "-- Choose Size/Model --" respectively when data validation was reset, ensuring the solution met the project's specific criteria.

Finally, to test the functionality and confirm my solution met all requirements, I saved the file as a .xlsm (macro-enabled) file and opened the grader file "Rental Center Inventory - GRADER.xlsm." I enabled macros, pressed the GRADE button, and if my solution was correct, I received a completion code for the project, which I could then use for the "Rental Center Inventory Submission Quiz."

Throughout the project, I made sure to maintain the integrity of the original inventory data, not removing any items, and only adding temporary data for testing purposes, which I promptly reset using the 'Original Data' tab when needed. This ensured that the project remained in line with the given guidelines and requirements.

## Distance Calculator

The task was to create an Excel dashboard that enables users to select starting and ending locations (states and cities) and calculates the straight-line distance between those two cities. Key requirements included data validation for state and city selections, dynamic city lists based on the selected state, unit selection for miles or kilometres, and displaying distance results. The project also involved using VBA scripts to reset conditional data validation and ensuring clear messaging for city and unit selection.

First, I set up data validation in cells B4 and B8 to restrict the user's selections to the 50 states and the District of Columbia. This was done using Excel's Data Validation feature.

Next, I implemented conditional data validation in cells B5 and B9. Depending on the state selected in cells B4 and B8, the drop-down list in cells B5 and B9 displayed only the cities within the selected state. This dynamic behaviour was achieved by creating named ranges for all 50 states and the District of Columbia and then using VBA scripts to reset the conditional data validation when needed. This ensured that if the user changed their state selection, the city options in cells B5 and B9 were updated accordingly.

I also formatted cells B5 and B9 to display "-- Choose City --" when cell B4 or B8 was changed, ensuring a clear indication for city selection. To achieve the formatting of cells B5 and B9 to display "-- Choose City --" when cell B4 or B8 was changed, I utilized Visual Basic for Applications (VBA). I created a VBA script that monitored changes in cells B4 and B8. When either of these cells was modified, the script automatically triggered an event that replaced the contents of cells B5 and B9 with "-- Choose City --" to provide a clear and consistent indicator for city selection. This dynamic updating ensured that users always saw this message when changing their state selections, guiding them to choose a city within the newly selected state.

In cell C11, I set up data validation to allow the user to choose between "mi" (miles) and "km" (kilometres) for distance units.

To calculate the straight-line distance between two cities, I referenced the latitude and longitude data provided in the "Data" worksheet and used the Haversine formula provided in the project description. This formula calculated the distance in miles or kilometres based on the user's selection in cell C11.

Finally, I ensured that cell B11 displayed "-- Choose City --" if either of the city selections in cells B5 or B9 was not made or if the selected cities were not contained within their respective states.

By following these steps, I created a functional Distance Calculator dashboard in Excel that met all the project requirements and allowed the user to select starting and ending locations to calculate straight-line distances between cities.

## Dinner Sign-Up

The given task was to create an Excel spreadsheet that automatically keeps and displays only the most recent submission for each participant who filled out a Google Form survey for a dinner party. This spreadsheet needed to be set up in a way that maintains alphabetical order by name and dynamically updates itself in real time if changes are made to the survey data.

To accomplish this, I utilized the INDEX and MATCH functions. The INDEX-MATCH combination allowed me to perform a reverse search starting at the last entry for each participant in the "Raw Data" worksheet. By carefully structuring the MATCH function to find the most recent entry for each participant, I then used the INDEX function to retrieve the corresponding information, such as Name, Item of food/beverage, and whether or not they will bring a Guest, from the "Raw Data" worksheet and display it in the "FINAL Selections" worksheet.

I ensured that the INDEX-MATCH formulas were applied correctly to the respective columns in the "FINAL Selections" worksheet, establishing a direct link between the data in this sheet and the corresponding data in the "Raw Data" worksheet.

Additionally, I made sure that my solution was designed to auto-update. Any changes made to the "Raw Data" worksheet would be automatically reflected in the "FINAL Selections" worksheet in real-time, meeting the project's dynamic update requirements.

In summary, I successfully solved the problem by using the INDEX and MATCH functions in Excel to automatically extract and display the most recent submissions for each participant in the "FINAL Selections" worksheet while ensuring that the data remained updated in real-time. This approach efficiently organized the dinner party sign-up data as required.

## Bakery Shopping List

To solve the Bakery Shopping List project, I followed these steps:

* Data Validation: I began by implementing data validation for cell E17 (the Name input cell) and cell E18 (the Ingredient input cell). For cell E17, I created a drop-down list with the names Abby, Bill, Cathy, Derek, and Emily, ensuring they were spelled correctly. For cell E18, I also set up a drop-down list with the 10 allowed ingredients: Flour, Sugar, Baking powder, Baking soda, Salt, Milk, Butter, Vanilla, Eggs, and Bananas.
* Cumulative Amount Calculation (Cell E20): To calculate the total cumulative amount of the selected ingredient in cell E18 ordered by the employee in cell E17, I used the SUMIF function. In cell E20, I applied the SUMIF function with E17 as the criteria range (employee name), E18 as the criteria (ingredient name), and the corresponding range of ingredient quantities. This calculated the total amount of the selected ingredient ordered by the employee.
* Unit of Measurement (Cell F20): To display the unit of measurement for the selected ingredient in cell E18, I used a formula that checked the selected ingredient and returned the appropriate unit. For example, if Flour was selected, the formula would display "cups." If Bananas or Eggs were selected, it displayed no unit, ensuring it wouldn't show "0" for those ingredients.
* Handling New Rows: I ensured that my solution could handle the addition of new rows to the table provided. The formulas and data validation rules were set up in a way that they automatically updated when a new record was added.

In summary, I implemented data validation for name and ingredient inputs, calculated the cumulative amount of the selected ingredient ordered by the employee, displayed the unit of measurement, and made sure the solution accommodated new rows. This approach met all the requirements of the Bakery Shopping List project.

## Historical Weather Lookup (Part A)

To solve the Historical Weather Lookup problem, I started by setting up data validation in Excel to create dropdown lists for the year, month, and day. This allowed users to conveniently select a date ranging from 01-Jan-1897 to 31-Dec-2020. Once the date was selected, I employed a combination of Excel functions to fetch the corresponding weather data from a large dataset.

To retrieve the high temperature, low temperature, precipitation, and snow values, I used the INDEX and MATCH functions. First, I constructed a unique list of dates from the dataset using the UNIQUE function, making sure it covered the entire date range. Then, I employed the TEXT function to format the selected date into the same format as the dataset.

With the date in the right format, I used the MATCH function to find its position in the unique list. This position was then used as the row number in the INDEX function to fetch the weather data for that specific date. To account for cases where the data wasn't available, I implemented an IF statement to return "not available" if no data was found.

In summary, I solved the Historical Weather Lookup problem by enabling user-friendly date selection, reformatting dates, creating a unique list of dates, and using INDEX, MATCH, and IF functions to retrieve the weather data or return "not available" when necessary. This approach allowed users to easily access historical weather information within the specified date range.

## Dynamic Data Cleaning

To address the issue of dynamic data cleaning within Excel, I first identified the specific column containing the data that needed to be cleaned. My primary goal was to eliminate blank spaces, strings, and values of 9999 from this column. To achieve this, I employed a series of Excel functions and formulas.

Initially, I utilized the "FILTER" function to isolate the data in the targeted column. By applying the "NOT" function in combination with the "ISBLANK" and the "ISNUMBER" functions, I was able to create a filtering condition. This condition allowed me to exclude blank cells and non-numeric values (i.e., strings) from the data. Additionally, I included a condition to exclude any cells with the value 9999.

By doing so, I effectively cleansed the data in the specified column, removing unwanted elements such as blank spaces, strings, and the value 9999. This dynamic data-cleaning process ensured that the column contained only the desired numeric values, ready for further analysis or utilization in various Excel functions and calculations.

## Friday the 13th

To solve the problem of calculating the total number of Friday the 13th occurrences within a span of two days using Excel, I employed a combination of Excel functions and formulas. Firstly, I created a date range starting from the beginning date using the SEQUENCE function, specifying the range of days within the two-day period. Then, I utilized the DAYS function to calculate the number of days between the start date and the generated date range.

To identify the days of the week, I used the WEEKDAY function on the generated date range. This function returns the day of the week as a number (1 for Friday, 2 for Saturday, and so on). By applying WEEKDAY, I could determine which days fell on a Friday (1).

Next, I used a conditional statement, to check if the day was equal to 5 (Friday) and if the date was the 13th of the month. When both conditions were met, I incremented a counter using the SUM function to keep track of the occurrences of Friday the 13th.

In summary, by employing Excel's SEQUENCE, DAYS, WEEKDAY, and SUM functions, I efficiently calculated the total number of Friday the 13th occurrences within the given two-day period, ensuring an accurate and automated solution to the problem.

## Sensitivity Analysis Problem Statement

To address the one-at-a-time (OAT) sensitivity analysis for our project, I employed Excel as a powerful tool to model the influence of various factors on the land's value. These factors encompassed the cost of land, royalties, total depreciable capital, working capital, start-up costs, sales, cost of sales, tax, and interest rates. To carry out this analysis effectively, I meticulously designed an Excel file that allowed me to observe how subtle alterations in these factors cascaded through the entire model, affecting not only individual values but also the overarching project valuation.

In Excel, I harnessed the "What-if Analysis" tool, a feature specifically designed for scenarios like this, where we need to explore the impact of changing variables on multiple interconnected calculations. This tool allowed me to systematically modify one factor at a time while keeping the others constant, facilitating a clear understanding of each factor's sensitivity and its contribution to the overall project dynamics. Through this methodical approach, I could assess how slight adjustments in any of these variables influenced not only the specific parameters but also the comprehensive financial picture. This data-driven analysis in Excel proved invaluable in making informed decisions, fine-tuning our project strategy, and optimizing our land investment.

## Dynamic Temperature Lookup

I began by implementing data validation in cells F5 and F6 using drop-down lists. I used the hours from column A as the options for these lists, ensuring they would update dynamically if more rows were added to the data. Then, I entered formulas into cells F8, F9, and F10 to calculate the average, maximum, and minimum temperatures, respectively. For the calculations, I employed the AVERAGEIFS, MAXIFS, and MINIFS functions, using criteria based on the selected starting and ending hours. To ensure the solution was dynamic, I converted the data into an Excel table. This allowed the temperature lookup cells to automatically adjust as new data was added starting from row 28. With these steps, I successfully created a solution that met the project's requirements for dynamic temperature lookup.

## Dynamic Amortization Schedule Problem Statement

I tackled the challenge of creating a dynamic amortization schedule that adjusts seamlessly to varying loan terms. Previously, the schedule was static, leading to problems when users changed loan terms. To address this, I harnessed the power of Office 365's dynamic array functions. I learned alternative approaches for setting up an amortization schedule and then transformed these methods into dynamic array formulas. As a result, the schedule now automatically adapts to the user-specified loan term, providing real-time updates.

## Longest Ladder Around The Corner

In this assignment, the task was to utilize Excel's Solver tool to determine the longest ladder that could fit around a corner formed by the intersection of two hallways. The hallways, with widths represented by variables w1 and w2, intersect at a right angle. By employing the provided formula, I established a relationship between the variables. Leveraging Solver, I adjusted the variable x to minimize the expression for the length of the ladder (L). This resulted in the maximum length of the ladder that could navigate the corner. Utilizing the "Longest Ladder Around the Corner.xlsm" file, I inputted hallway widths, engaged the Solver tool, and obtained the desired ladder length. Finally, I ensured the solution's robustness by re-evaluating it when inputs changed.

## World Bank Lookup

In the completion of the World Bank Lookup assignment, I engaged with a 2009 database featuring 16 socio-economic indicators across 171 countries. By downloading the provided "World Bank Lookup.xlsm" file and enabling macros, I established a functional dashboard. Utilizing data validation techniques, I implemented dropdown lists for country selection in cell C3 and socio-economic indicators in cell C4. This allowed me to employ the "LOOKUP" function to locate corresponding data on the "Data" tab and display it accurately in cell C6 of the "Dashboard" tab. I ensured adherence to the specified criteria and drew insights from provided hints and references, ultimately achieving a comprehensive and operational solution.

Through this endeavour, I gained proficiency in the strategic implementation of Excel functions such as "LOOKUP" and data validation protocols, allowing for the creation and management of highly functional dropdown lists tailored to specific informational requirements.

## Nearest Eight of an Inch

To complete this assignment, I commenced by receiving a decimal-length measurement, inputted in cell B4 of the macro-enabled file "Nearest Eighth of an Inch.xlsm". The objective entailed transforming this measurement into its corresponding values in feet, inches, and the nearest 1/8th of an inch, each to be displayed in cells B7, B8, and B9, respectively.

For instance, when dealing with a measurement of 7.4 feet, my approach was as follows:

1. Isolate the integer portion, yielding 7 feet.
2. Deduct this integer part from the original measurement (7.4 - 7) to determine the residual fraction of a foot, which in this case was 0.4.
3. Given that 12 inches constitute a foot, I calculated the equivalent inches by multiplying the fraction (0.4) by 12, resulting in 4.8 inches.
4. I repeated a similar procedure as in step 2 to ascertain the fraction of an inch, which in this scenario was (4.8 - 4) = 0.8 inches.
5. Lastly, I rounded this fractional inch value (0.8) to the nearest 1/8th of an inch, where 1/8th equals 0.125 inches.

It is important to note that the assignment's guidelines, along with provided resources like the screencast "Nearest Eighth of an Inch preview," and references to relevant "Everyday Excel" sections, substantially facilitated the completion of this task. The solution automatically updated in response to changes in the measurement within cell B4.
