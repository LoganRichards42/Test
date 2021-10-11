import requests as api
import json

from time import sleep
import xlsxwriter

searchText = input("Enter site name / group name: ")
api_key = 'SOYBB735OIGU2V32M3Q3P5QEZ89FSW0U'
url = 'https://monitoringapi.solaredge.com/sites/list.json'
params = {'api_key':api_key,'status':"Active",'searchText':searchText}

Orange = '#fce4d6'
Blue = '#ddebf7'

Makats = {
  10:317887994018, # Makat doesn't exist, using 12.5 makat instead 
  12.5:317887994018,
  16:317887994007,
  17:317887994006,
  25:317887994005,
  27.6:317887994004,
  33.3:317887994004, # Not real Makat since none exists
  50:317887994003,
  55:317887994014,
  60:317887998030,
  66.6:317887998030,
  75:317887994013,
  82.8:317887994002,
  100:317887994017
}

def FindMakat(InverterModel,Location):
  for key in Makats.keys():
    if (InverterModel == float(key)):
      worksheet.write('B'+Location,str(Makats[key]),cell_format) # Formatting only
      break
    else:
      worksheet.write_blank ('B'+Location, '',cell_format) # Formatting only

def FindInverterModel(InverterModel):
  InverterModel = str(InverterModel)
  InverterModel = InverterModel.split("K-")[0]
  InverterModel = InverterModel.split("SE")[1]
  InverterModel = float(InverterModel)
  return InverterModel


workbook = xlsxwriter.Workbook('SolarEdgeToPriority.xlsx',{'constant_memory': True})

header_format = workbook.add_format()
header_format.set_font_name('Calibri')
header_format.set_font_size(10)
header_format.set_text_wrap()
header_format.set_align('vcenter')
header_format.set_align('center')
header_format.set_border()

top_cell_format = workbook.add_format()
top_cell_format.set_border()
top_cell_format.set_align('vcenter')
top_cell_format.set_align('center')

bottom_cell_format = workbook.add_format()
bottom_cell_format.set_border()
bottom_cell_format.set_align('vcenter')
bottom_cell_format.set_align('center')

worksheet = workbook.add_worksheet()


worksheet.write('A1','מספר הלקוח בפריוריטי',header_format)
worksheet.write('B1','מק"ט',header_format)
worksheet.write('C1','מספר מכשיר',header_format)
worksheet.write('D1','מספר אתר בפריוריטי',header_format)
worksheet.write('E1','קוד אחריות',header_format)
worksheet.write('F1','אחוז חריגה לממיר',header_format)
worksheet.write('G1','הספק מכשיר ביח KV [AC]',header_format)
worksheet.write('H1','הספק מקסימלי למכשיר ביח KV [DC]',header_format)
worksheet.write('I1','מספר אופטימייזרים לממיר',header_format)
worksheet.write('J1','מס"ד לקוח',header_format)
worksheet.write('K1',"מספר טבוע",header_format)
worksheet.write('L1','שם אתר בפורטל',header_format)

numSiteReadInLastCall = -1 
numSitesRead = 0 
sites = []

Location = 2
MaxWidth = 0
Color = True

# Loop below gets all sites
for i in range (0,100):
  params['startIndex'] = numSitesRead
  Answer = api.get(url, params = params)
  siteList = json.loads(Answer.text)
  siteList = siteList['sites']['site']
  numSiteReadInLastCall = len(siteList)
  numSitesRead += numSiteReadInLastCall
  sites.extend(siteList)
  if numSiteReadInLastCall == 0:
    break
# Done, all sites stored on "sites: var

t = json.dumps(sites, indent = 4, sort_keys=True,ensure_ascii=False)
T = json.loads(t)


for x in range(len(T)):

  tempid = T[x]['id'] # Site ID 
  tempname = T[x]['name'] # Site name
  peakPower = T[x]['peakPower']

  try:
    moduleModel = T[x]['primaryModule']['modelName']
    moduleWatt = T[x]['primaryModule']['maximumPower']
  except:
    print(str(tempid) + "  -  " + tempname[::-1] + " has no module")

  if len(tempname) > MaxWidth:
    MaxWidth = len(tempname) + 1

  url = 'https://monitoringapi.solaredge.com/site/'+str(tempid)+'/inventory.json'
  params = {'api_key':api_key}

  sleep(0.1)
  Answer = api.get(url, params=params)
  Inverters = json.loads(Answer.text)
  Inverters = Inverters['Inventory']['inverters']
  Color = not Color
  xC = 1
  for I in Inverters:
    cell_format = workbook.add_format()
    cell_format.set_border()
    if (xC == 1):
      cell_format.set_top(2)
    elif (xC == len(Inverters)):
      cell_format.set_bottom(2)

    xC+= 1

    cell_format.set_align('vcenter')
    cell_format.set_align('center') 

    if Color == True:
      cell_format.set_bg_color(Blue)
    else:
      cell_format.set_bg_color(Orange)

    Location = str(Location)

    worksheet.write_blank ('A'+Location, '',cell_format) # Formatting only
    worksheet.write_blank ('D'+Location, '',cell_format) # Formatting only
    worksheet.write_number('E'+Location,0,cell_format)
    worksheet.write_number('F'+Location,15,cell_format)
    worksheet.write('L'+Location, tempname,cell_format)
    worksheet.write('J'+Location, tempid,cell_format)

    inverterDC = (int(moduleWatt) * 2 * int(I['connectedOptimizers'])) / 1000
    worksheet.write('H'+Location, inverterDC, cell_format)

    
    try:
      worksheet.write('I'+Location, I['connectedOptimizers'], cell_format)
      worksheet.write('C'+Location, I['SN'], cell_format)
      
      InverterModel = FindInverterModel(I['model'])

      worksheet.write_number('G'+Location,InverterModel, cell_format)  # Inverter AC
      worksheet.write('K'+Location, I['name'], cell_format)

      FindMakat(InverterModel,Location)

    except Exception as E:
      print(tempname[::-1] + I['name']+' out of ' + str(len(Inverters)) + " failed "+ str(E) +"\n")
    Location = int(Location) + 1
  print(str(x+1) + ' out of '+ str(len(T)) + "\n")

my_format = workbook.add_format()
my_format.set_align('vcenter')
my_format.set_align('center')
worksheet.set_column('L:L', MaxWidth)
worksheet.set_column('B:C', 12.5)
worksheet.set_column('A:R', None, my_format)
worksheet.right_to_left()

workbook.close()
print("\nDone!")
