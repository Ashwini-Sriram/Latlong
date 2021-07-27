from opencage.geocoder import OpenCageGeocode
import xlrd
import xlwt
from xlwt import Workbook
import pandas as pd

key ="fd4f682cf2014f3fbd321ab141454138" 
# get api key from:  https://opencagedata.com


	
geocoder = OpenCageGeocode(key)


	



loc = ("/Users/ashwinisriram/Documents/Lat long/corrected.xlsx")

wb = xlrd.open_workbook(loc)
sheet = wb.sheet_by_index(0)
sheet.cell_value(0, 0)

# Workbook is created
wb = Workbook()

# add_sheet is used to create sheet.
sheet1 = wb.add_sheet('Sheet 1')

#  Define a dictionary containing  data
data={'Customer_code':[],'City':[],'State':[]}
branch_district = ""
  

for r in range(4000,4500):
    customer_code=str(sheet.cell_value(r, 0))
    #   sheet1.write(i, 1, sheet.cell_value(r, 1))
    #   sheet1.write(i, 2, sheet.cell_value(r, 2))
    branch = str(sheet.cell_value(r, 1))
    district= str(sheet.cell_value(r, 2))
    data['Customer_code'].append(customer_code)
    data['City'].append(branch)
    data['State'].append(district)
    


df=pd.DataFrame(data)


# Convert the dictionary into DataFrame


# Observe the result
print(df)




list_lat = []   # create empty lists

list_long = []

link=[]

for index, row in df.iterrows(): # iterate over rows in dataframe



    City = row['City']
    State = row['State']       
    query = str(City)+','+str(State)

    results = geocoder.geocode(query) 
    try:
        lat = results[0]['geometry']['lat']
        long = results[0]['geometry']['lng']
        list_lat.append(lat)
        list_long.append(long)
        link.append("http://www.google.com/maps/place/"+str(lat)+','+str(long))


    except IndexError:
        list_lat.append('Nan')
        list_long.append('Nan')
        link.append("link unavailable")


    
	

# create new columns from lists    

df['lat'] = list_lat   

df['lon'] = list_long

df['link']=link
# function to find the coordinate
# of a given city
print(df)
# create excel writer object
writer = pd.ExcelWriter('output2.xlsx')
# write dataframe to excel
df.to_excel(writer,'sheet2')
# save the excel
writer.save()
print('DataFrame is written successfully to Excel File.')

