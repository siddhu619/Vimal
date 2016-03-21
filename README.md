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
