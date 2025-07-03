# QR_Ticket-generator
This app creates tickets with a qr-code for text-output.

You need:

* An png image file for the tickets. Format should be 4:3.
* A xlsx or csv list of participants. Columns should be 1: ID/Nr, 2: Date, 3: Name and 4: Ticket category.
* A thought on what text to be displayed when QR-code is scanned. The following codes can be placed inside the text: {date} = date, {ticketnr} = ID/nr, {name} = name, {ticketcat} = ticket category and {nl} = new line.

Example text: Spring Concert: Local heroes{nl}{nl}Registration date:{date}{nl}Ticketnr.: {ticketnr}{nl}Name: {name}{nl}Ticket category: {ticketcat}

Exports are:

* A folder called "QR-codes", containing all QR-codes as image files
* A folder called "Tickets", containing all tickets as pdf-files
* An Excel-file called "Orders_list" with all participants, and a column to make comments, i.e. mark at arrival.

IMPORTANT: 
To read the QR-code not all QR-readers can be used. Some readers will convert the text to an internet search. 
One reader which is usable for reading this apps QR-codes is "QR Reader for iPhone" by TapMedia Ltd.
