# Email Address Counter

This Python program connects to an email server, retrieves a list of all emails in the inbox, and counts the number of unique .edu email addresses. The program then creates an Excel file with the email addresses and their respective counts. It also extracts the domain names from the email addresses and counts the number of times each domain appears in the list. It then creates hyperlinks to the unique domains and writes them to a new column in the same Excel file. Finally, it removes duplicate email addresses and domain names.
Installation

### Clone the repository:

    git clone https://github.com/username/repository.git

### Install the required packages:

    pip install -r requirements.txt

## Usage

Change the email and app-specific password in the script to your own email and password. 
You can generate an app-specific password by following these instructions: [here](https://myaccount.google.com/u/0/apppasswords?rapt=AEjHL4PoZPQ7LovSpkNjHn4Yp2UVlegzTah7sTWHVQM3wydBn8tDtINwbduObbIlvfolo3KzcE2v5qXqpviUv0g7wZTHfX5qDw&pageId=none&pli=1.)

### Run the script:

    python email_counter.py

The program will generate an excell xlsx file.
