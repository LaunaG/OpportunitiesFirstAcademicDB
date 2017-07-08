# Opportunities First Academic Database

**CREATED:** 2016

**LANUGAGE:** Visual Basic for Applications

**STATUS:** IN PROGRESS

**OVERVIEW:** A local nonprofit running a WIOA (Workforce Innovation and Opportunity Act) program for youth aged 17-24 did not have a scheme for storing, retrieving, visualizing, and analyzing their clients' test scores to gauge the effectiveness of their educational programs.  Assessments included: the TABE (Test of Adult Basic Education) Locator and Survey, Official GED Practice tests, Official GED tests, OhioMeansJobs practice WorkKeys tests, and the official ACT WorkKeys tests. To address this need, I began developing a Microsoft Access tool that allowed for the entry and modificiation of student demographic and test score data and provided an interactive calendar to track scheduled GED tests.

**NEXT STEPS IN PROJECT:**
* Refactor and standardize variable names following completion of coding boot camp
* Possibly rewrite the program as a dynamic web application using the .NET framework and MS SQL Server to ensure better scaling and maintenance
* Add functionality to the "Help" buttons
* Determine the best location to modify existing test score data
* Finish the last section, "Data Analysis", by (1) writing queries to summarize student test scores with descriptive statistics, (2) creating graphs to display the statistics results, and (3) enabling users to perform inferential tests on the data
* Add column for student age and gender
* Add OGT (Ohio Graduation Test), ServSafe, Customer Service, and Hospitality tests as tables in the database and include them in the score entry form
* Create unit tests with sample data

#### MAIN SCREEN
The admin chooses from among four different options: (1) modifying cohort information, (2) entering test score data, (3) viewing or creating upcoming test appointments with an interactive calendar, and (4) summarizing and analyzing test data at the individual student, cohort, or multiple cohort level. *NOTE: The "Data Analysis" option is not functional at this time.*

![alt text](https://github.com/LaunaG/OpportunitiesFirstAcademicDB/blob/master/homePage.png "Home Page")

#### STUDENT INFORMATION
Students enter the nonprofit's program in groups called cohorts.  Here, admins can create or modify cohort names and then add students to each cohort.  Available fields for students are currently first name, last name, and initial diploma status. Admins can return to the main screen by clicking the "Home" button.  In addition, on every page, they can click a "Help" button (question mark) to receive instructions for navigating the content there. *NOTE: The "Help" buttons are not functional at this time.*

![alt text](https://github.com/LaunaG/OpportunitiesFirstAcademicDB/blob/master/Cohorts.gif "Edit Cohorts gif")

#### SCORE REPORTING
Admins can enter new test scores by using the tab control to switch between tests. To prevent database corruption, several input fields have a layer of data validation; in addition, when the form is submitted with blank required fields, the program will respond with an error message and return to the incomplete form.

![alt text](https://github.com/LaunaG/OpportunitiesFirstAcademicDB/blob/master/ScoreEntryLowerFrame.gif)

#### TEST SCHEDULING
When the form opens, admins land on a calendar with the current month displayed and the current day outlined in green.  Previous days are greyed out, while current and upcoming days are left white.  Admins can change the month and year by using the drop-down boxes and can iterate to the previous or next month using the left and right arrows, respectively.  Each calendar box displays the number of tests scheduled for that day; when clicked, a modal window opens and displays more detailed information (i.e. name, test date, test time, test subject, location, transportation, comments). Admins can also schedule a new test by clicking on the "New Test Appointment" button and filling out the modal form that launches.

![alt text](https://github.com/LaunaG/OpportunitiesFirstAcademicDB/blob/master/CalendarLowerFrame.gif "Calendar gif")
