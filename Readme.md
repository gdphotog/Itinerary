This app is for creating a basic Itinerary to enable the IQWF system to stay running when the ships are in lockdown.

May 29,2021 - Update 1.0.0.5
Config file has been added to the project.
    3 options in the config file:
        DebugON - This will enable the debugging features added into the app which will show the message boxes at certain points of the system
        NumberOfGuests - This is a numeric value that will give the manifest length
        RunMode - This has 2 values (LAB or SHIP) In Lab mode it will create the manifest, Itinerary and Staff files whilst in Ship mode it will only create Itinerary and Staff files.
    Hotfixes -
        Random number generation was returning the same number each loop due to the way VB.net generates them based on time. 

October 6, 2021

RunMode - added HOME as a value to have some of the RFID cards utilized from my stack here.
