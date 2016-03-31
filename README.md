# Vimal

**Flow of script:**

1. Fetches 1000 records from excel file.
2. Checks validation on the fetched data.
3. Removes non-ascii character present in address field.
4. Generates json and sends it to the service api.
5. Reads the response from the server and writes it to a file  in the format below:

    "Status"  : "Reason": "EmailId of record"  
    e.g: failure : Error Inserting email : nitinghope@yahoo.com
6. After writting the file for 1000 responses, It reads the file and checks for the  response have status as "failure"

    if the number of records have failed status is more than one it stops the script (you should check the file 
      and find why it happend)
      
    if the number of records which have failed status is 0, the script will again fetch next 1000 records, parses 
      and sends it to the server (infinite loop until one records fails or all records are fetched)
      
      
7. Change the file paths according to your system before executing it.

**Latest Changes Added:**

1. Removes non-ascii as well as quotes and escape character which used to stop the execution of code.
2. New Feature Multiprocessing added to send data to server in multi processing manner
    
    Currently we are using 15 simulataneous processes to acomplish our job which I find is optimized.
    
    If we want you can increase it but it depends upon the number of Cores you have in your system and How much the 
        server can handle.
    
    Currently is process requires 50MB of RAM so manage your memory accordingly
3. After sending the data to the server the codes writes the response as well as the record for which it send in the
    below format:
        "User Record essential data" $$$$$$$$$ "Response from server"
4. I am planning to write another script which excuted after the above code executes to read all the user records  
    and connect to DB to verify if the data has been inserted properly
    
    To accomplish the above task I need you DB details as well as crendentail.


