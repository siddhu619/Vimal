from openpyxl import load_workbook

import requests

print 'Inserting the the data collected from Database'
wb = load_workbook('C:\\maharastra\\mumbai.xlsx',data_only=True)
worksheet = wb.get_sheet_by_name("result_ok_7011_2016-03-03 (1)")
#worksheet = wb.get_sheet_by_name("Sheet1")
print 'Fetching the test data from database '

name = []
phone= []
email= []
address= []
dob= []
city= []
edu = []
gender = []




for i in range(1, 1000):

    name.append(worksheet['A'+str(i)].value)
    email.append(worksheet['B'+str(i)].value)
    dob.append(worksheet['D'+str(i)].value)
    phone.append(worksheet['E'+str(i)].value)
    city.append(worksheet['P'+str(i)].value)
    address.append(worksheet['AC'+str(i)].value)
    #print str(worksheet['A'+str(i)].value)
    edu.append(worksheet['R'+str(i)].value)
    gender.append(worksheet['AA'+str(i)].value)
Fname= []
Lname = []
#print name[0]
for n in range(999):
    #print n
    phone[n] = str(phone[n]).replace(" ",'').replace('-','')
    d = name[n].split(' ')
    Fname.append( d[0])
    Lname.append(d[len(d) -1])

with open('E:\\ResultJson.txt','w')as f:
    for j in range(999):


        regDetails= '{"birthday":"'+str(dob[j]).replace('-','/')+'", "lastName":"'+ str(Lname[j])+'", "firstName": "'+str(Fname[j])+'","password":"ZO25L2", "email_address" : "'+str(email[j])+'","education": "'+str(edu[j])+'","gender":"'+str(gender[j])+'","referredBy": "","phone_number": "'+ phone[j]+'","newFlow": true' +'}'

        addressDetails= '{"country_code": 356, "is_primary": true, "city":"'+str(city[j])+'", "address_1": "'+str(address[j])+'","address_type_id":"1", "state_code":3880,"zip_code": "400096"}'

        regDetailparam = '{"extraGoogleParams": {} ,"addressDetails"  :'+ str(addressDetails)+ ', "RegistrationDetails":  '+ str(regDetails) +'}'
        print regDetailparam
        with requests.Session() as s:
            pass
            #r = s.post('https://perf.reflexisinc.com/PULSEV31/systemAction.htm', data=login_data1,verify=False)
            m = s.post("https://app.ship2myid.com/ship2myid/client/rest/webclient/autoregister_full?reg_source=email",data=regDetailparam,headers = {"Accept": "application/json", "Content-type":"application/json"}, verify=False)
            print m.status_code
            print m.text

            f.write(m.text)