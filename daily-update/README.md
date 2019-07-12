Daily Update Emails
==========

A collection of files in Google Apps Scripts used to collate data from various sources and sends this as an email.

The data collected is as follows:

- Calendar events (calendar.gs) which checks multiple Google calendars for all-day events (can be adapted to check a range of times)
- Weather data (weather.gs) which checks the humidity for the upcoming evening at 5 timepoints between 6pm and 6am to alert to >90% humidity
- Waste collections (waste.gs) which holds information on the upcoming household waste and recycling collections

These are called by the main file (main.gs) which collates and sends the email, and can be set to trigger once/twice per day.

Feel free to use/adapt as required and leave me any comments or feedback!

Thanks,

Matt
