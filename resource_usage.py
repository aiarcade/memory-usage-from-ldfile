import sys

from openpyxl import Workbook
from openpyxl import load_workbook
from openpyxl.styles import Font, Color
from openpyxl.styles import colors

class Map():
    def __init__(self,line,stype):
        self.type=stype
        valid_data=filter(None,line.split(" "))
        self.section=valid_data[0].replace("\t","")
        self.address=valid_data[1].split("\t")[0]
        self.size=int(valid_data[1].split("\t")[1],16)
        if self.type=='obj':
            self.file_name=valid_data[-1].split("\\obj\\")[1].replace("\r","").replace("\n","")
        else:
            self.file_name=valid_data[-1].split("\\lib\\")[1].replace("\r","").replace("\n","") 
        print self.section,self.address,self.size,self.file_name
            
        

input_file=sys.argv[1]
output_file=sys.argv[2]

map_file=open(input_file,"r")
data=map_file.readlines()

start_map=0
maps=[]

for line in data:
   
    if line.find("Link Editor Memory Map")>-1:
        start_map=1
    if line.find(".debug")>-1:
        start_map=0
    if start_map==1:
        if line.find("\\obj\\")>-1:
            c_map=Map(line,'obj')
            maps.append(c_map)
            continue
        if line.find("\\lib\\")>-1:
            _map=Map(line,'lib')
            maps.append(_map)
            continue
        



index=0
ram_usage=0
rom_usage=0


existing_out_file=False
try:
    wb = load_workbook(output_file)
    existing_out_file=True
    ws_summary=wb.get_sheet_by_name("Summary")
except:
    wb = Workbook()
    ws_summary = wb.active
    ws_summary.title="Summary"

ws_data = wb.create_sheet()
ws_data.title=input_file
i=1
for _map in maps:
        d = ws_data.cell(row = i, column = 1)
        d.value=_map.section
        d = ws_data.cell(row = i, column = 2)
        d.value=_map.address
        d = ws_data.cell(row = i, column = 3)
        d.value=_map.size
        d = ws_data.cell(row = i, column = 4)
        d.value=_map.file_name

        i=i+1
        if _map.address[0]=='4':
            ram_usage=ram_usage+_map.size
        else:
            rom_usage=rom_usage+_map.size


s_r=ws_summary.get_highest_row()+3
s_c=ws_summary.get_highest_column()

summary_title_font= Font(color=colors.RED)
ws_summary.cell(row = s_r-1, column = 5).value=input_file
ws_summary.cell(row = s_r-1, column = 5).font=summary_title_font

ws_summary.cell(row = s_r, column = 5).value="Resource"
ws_summary.cell(row = s_r, column = 6).value="Bytes"
ws_summary.cell(row = s_r, column = 7).value="KBytes"


ws_summary.cell(row = s_r+1, column = 5).value="RAM"
ws_summary.cell(row = s_r+1, column = 6).value=ram_usage
ws_summary.cell(row = s_r+1, column = 7).value=ram_usage/1024



ws_summary.cell(row = s_r+2, column = 5).value="ROM"
ws_summary.cell(row = s_r+2, column = 6).value=rom_usage
ws_summary.cell(row = s_r+2, column = 7).value=rom_usage/1024
ws_summary.cell(row = s_r+2, column = 8).value="MAX KB"

    
       
        

wb.save(output_file)





