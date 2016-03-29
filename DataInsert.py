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
worksheet = wb.get_sheet_by_name("result_ok_7011_2016-03-03 (1)")
counter =  worksheet.get_highest_row()
flag =0





def RemoveNonAscii(address): # Removes non-asci character from the string and replaces it with blank string

    return ''.join([k if ord(k) < 128 else '' for k in address])



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
    ReadCount =1






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
    #print 'jgfsdfghjkjhv',start, end
    #print ReadCount
    #print name[0]
    #print 'sadfsdfsdfsdfsdsdfsdf',len(email)
    for n in range(len(email)):
        if name[n] == None:
            name[n]= ''

        elif  phone[n] == None:
            phone[n] = ''
            #print type(phone[n])
        elif isinstance(address[n], (int,long)):
            address[n] =''


        phone[n] = str(phone[n]).replace(" ",'').replace('-','')
        #print len(name[n].split(' '))
        #print name[n]
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

def PostData(regDetailparam):
    with open('E:\\ResultJson.txt','a')as f:
        #print len(regDetailparam)
        print regDetailparam

        #f.truncate()
        with requests.Session() as s:
            m = s.post("https://app.ship2myid.com/ship2myid/client/rest/webclient/autoregister_full?reg_source=email",data=regDetailparam,headers = {"Accept": "application/json", "Content-type":"application/json"}, verify=False)
            print m.status_code, datetime.datetime.now()
            print m.text

            resp = json.loads(str(m.text))
            #f.write(resp['Error.status'] +' : '+ resp['reason']+' : '+email[j]+'\n')
            f.write(json.loads(str(regDetailparam))['RegistrationDetails'])

    f.close()



def main():
    global counter

    if ReadCount > counter:
        exit()
    #print 'sadfsfsdf',ReadCount
    global flag

    if flag ==0:
        Read1000(ReadCount ,ReadCount + 999)
        flag =1


    else:
        Read1000(ReadCount,ReadCount + 1000)
    with open('E:\\ResultJson.txt','a')as f:
        #f.truncate()
        regDetailparam =[]


        for j in range(len(email)):

            #print 'asdfsdfsdfsdfs',j
            regDetails= '{"birthday":"'+str(dob[j]).replace('-','/')+'", "lastName":"'+ str(Lname[j])+'", "firstName": "'+str(Fname[j])+'","password":"ZO25L2", "email_address" : "'+str(email[j])+'","education": "'+str(edu[j])+'","gender":"'+str(gender[j])+'","referredBy": "","phone_number": "'+ str(phone[j])+'","newFlow": true' +'}'

            addressDetails= '{"country_code": 356, "is_primary": true, "city":"'+str(city[j])+'", "address_1": "'+RemoveNonAscii(address[j])+'","address_type_id":"1", "state_code":3880,"zip_code": "400096"}'

            regDetailparam.append('{"extraGoogleParams": {} ,"addressDetails"  :'+ addressDetails+ ', "RegistrationDetails":  '+ str(regDetails) +'}')
            #print regDetailparam

        p = Pool(5)
        p.map(PostData,regDetailparam)

                #f.write(str(j)+'email[j]'+email[j]+'\n')
    f.close()


    '''
    with open('E:\\ResultJson.txt','r')as f1:
        FailedCount = 0
        for line in f1:

            if line.find('failure') >= 0:
                print line
                FailedCount =+1
                #print FailedCount
        print FailedCount
        #print 'asdfsdfsdf',ReadCount
        '''
    #main()
    '''
        if FailedCount == 0:
            main()
        else:
            exit()
        '''




if __name__ == '__main__':
    main()
