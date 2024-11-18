# D365 Dashboard Project
A combined web resource dashboard to implement into D365 to allow for a centralised and UX friendly place to view information.

## Table  of Contents
- [Usage](#usage)
- [Contributors](#contributors)

## Usage
### HTML
    1a. Tabs - Tabs for the side nav are located under each div as "sidebar", the active "page" as it were, is nominated by the class of "active".
    1b. Body - the body of each div is determined by various format types however the preset types are infoGrid and fieldTable, each has preset CSS to determine the styling.
### CSS
    2a. Styling - Styling for this piece is completely customisable however, I have gone with my company colours and kept it roughly in keeping with the UI of D365
### JavaScript
    3a - Functions - the main function is firstLoad(), this loads all the fields present in our CRM, these can be customised to your liking, I have created multiple functions to work regardless of the environment and should be pretty simple to reuse.
    3b - Some functions are hard coded into the JavaScript and will be refined and anonymised so that they can be used better outside of my current CRM system.

## Contributors
[Jake Steele](https://github.com/jakesteeledev/)