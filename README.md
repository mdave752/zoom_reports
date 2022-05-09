# zoom_reports
Creates a readble excel report as .xlsx file, from zoom meeting detail csv.
Had a problem that School I worked at, had to make monthly reports to about attendence of youth.

<img width="865" alt="Screenshot 2022-05-09 at 19 36 06" src="https://user-images.githubusercontent.com/80170874/167456371-88dad3e6-7ce2-4307-a445-fcd97fe258f1.png">

Small region support:
    change REGION_PREFERENCES dictionary vaues to support your langauge
    change REGION_PREFERENCES["dateToLocal"] with your region specific date format using datetime strftime library
    posibility to filter by Unique names or Emails(In possibility if Email address is not required, can be a problem with multiple same names)
