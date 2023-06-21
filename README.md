#Hydrogen-Calculcator


We load the data for building tables from a document that is stored on Google Drive, for this we use the SheetJs library. While the data has not been uploaded, the loader will be displayed and the actions on the screen should be blocked.
After the data is loaded, we process it. We wrote our own data parser because the library did not process some Excel formulas. To calculate formulas, we use the jQuery Calx plugin (a jQuery plugin for creating formula-based calculation form). To build charts, we use the Chart.js library, the data is taken from the tables that we have hidden since we do not edit our Google document when recalculating. When the page is loaded for the first time, the table and chart data are taken from the document data stored on Google Drive, and when recalculated, the data is saved on the client side.
jQuery Tokenize is a plugin which allows your users to select multiple items from a predefined list or ajax, using autocompletion as they type to find each item. You may have seen a similar type of text entry when filling in the recipients field sending messages on facebook or tags on tumblr.

