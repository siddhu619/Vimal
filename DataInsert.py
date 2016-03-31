from openpyxl import load_workbook

import requests
import json
import datetime
from multiprocessing import Pool
import itertools
global ReadCount
ReadCount =1
global counter
wb = load_workbook('F:\\mumbai.xlsx',data_only=True)
worksheet = wb.get_sheet_by_name("Sheet1")
counter =  worksheet.get_highest_row()
flag =0
# truncates the result for fresh edit
global Response
Response = []
global f1
f1 = open('E:\\ResultJson.txt','a')



#!# Removes quotes, excape character and non-ascii characters from the address field

def RemoveNonAscii(address): # Removes non-asci character from the string and replaces it with blank string
    address= address.replace('"','')
    address= address.replace('\\','-')
    return ''.join([k if ord(k) < 128 else '' for k in address])


#@# Read 1000 records from the file at a time
def Read1000(start , end): # Reading 1000 lines form excel file
    print 'Collecting the data  from Excel'





    global ReadCount
    global name
    global phone
    global email
    global address
    global dob
    global city
    global edu
    global gender
    global Fname
    global Lname


    name = []
    phone= []
    email= []
    address= []
    dob= []
    city= []
    edu = []
    gender = []
    Fname= []
    Lname = []







    for i in range(start, end):



        name.append(worksheet['A'+str(i)].value)
        email.append(worksheet['B'+str(i)].value)
        dob.append(worksheet['D'+str(i)].value)
        phone.append(worksheet['E'+str(i)].value)
        city.append(worksheet['P'+str(i)].value)
        address.append(worksheet['AC'+str(i)].value)
        #print str(worksheet['A'+str(i)].value)
        edu.append(worksheet['R'+str(i)].value)
        gender.append(worksheet['AA'+str(i)].value)
        #print worksheet['B'+str(i)].value

        ReadCount = ReadCount +1
        if ReadCount > counter:
            print 'ReadCount',ReadCount
            break

    for n in range(len(email)):

        if name[n] == None:
            name[n]= ''

        elif  phone[n] == None:
            phone[n] = ''
            #print type(phone[n])
        elif isinstance(address[n], (int,long)):
            address[n] =''


        phone[n] = str(phone[n]).replace(" ",'').replace('-','')

        if len(name[n].split(' ')) >= 2:
            d = name[n].split(' ')
            #print d[0]
            Fname.append(d[0])
            Lname.append(d[len(d) -1])
            #print n,'------>',Fname[n]
        else:
            Fname.append(name[n])
            Lname.append('')

    name = []


#!# Muliprocessingly sends data to service/server
def PostData(regDetailparam):
        global Response

        print regDetailparam
        with requests.Session() as s:

            m = s.post("https://app.ship2myid.com/ship2myid/client/rest/webclient/autoregister_full?reg_source=email",data=regDetailparam,headers = {"Accept": "application/json", "Content-type":"application/json"}, verify=False)
            print m.status_code, datetime.datetime.now()
            print m.text


            Response.append(str(json.loads(str(regDetailparam))['RegistrationDetails'])+'$$$$$$$$$$'+str( m.text))
            print 'sjsjfjsfsdj'
            print len(Response),Response
            return str(json.loads(str(regDetailparam))['RegistrationDetails'])+'$$$$$$$$$$'+str( m.text)

#@# Method to write Response in a file
def ResponseWrite(Response):

    print '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!'
    print len(Response),Response
    with open('E:\\ResultJson.txt','a')as f:
        for r in Response:
            print '*****************************************************************************'
            rsp = r
            print 'Writing Reponse to file '+rsp
            f.write(rsp)
            f.write('\n')
    Response = []
    f.close()





def main():
    global counter

    #@# ReadCount is the No of records processed, counter is the maximum no record present in the file

    if ReadCount > counter:
        exit()

    global flag
    #@# Reading 1000 records from the file and processing it
    if flag ==0:
        Read1000(ReadCount ,ReadCount + 999)
        flag =1


    else:
        Read1000(ReadCount,ReadCount + 1000)
    with open('E:\\ResultJson.txt','a')as f:

        regDetailparam =[]


        for j in range(len(email)):


            ### Making List of 1000 json to be posted to server at a time
            regDetails= '{"birthday":"'+str(dob[j]).replace('-','/')+'", "lastName":"'+ str(Lname[j])+'", "firstName": "'+str(Fname[j])+'","password":"ZO25L2", "email_address" : "'+str(email[j])+'","education": "'+str(edu[j])+'","gender":"'+str(gender[j])+'","referredBy": "","phone_number": "'+ str(phone[j])+'","newFlow": true' +'}'

            addressDetails= '{"country_code": 356, "is_primary": true, "city":"'+str(city[j])+'", "address_1": "'+RemoveNonAscii(address[j])+'","address_type_id":"1", "state_code":3880,"zip_code": "400096"}'

            regDetailparam1 = '{"extraGoogleParams": {} ,"addressDetails"  :'+ addressDetails+ ', "RegistrationDetails":  '+ str(regDetails) +'}'
            regDetailparam.append(regDetailparam1)


        #Mention the number of process to post data on the service
        p = Pool(15)
        list = p.map(PostData,regDetailparam)
        ResponseWrite(list)




    f.close()

# Recurssive call of main untill all records are processed!!
    main()




if __name__ == '__main__':
    main()
