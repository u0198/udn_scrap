#!/usr/bin/env python
# coding: utf-8

# In[1]:


from datetime import datetime, timedelta
import win32com.client as client

excelapp = client.Dispatch("Excel.Application")
excelapp.Visible = True

date = datetime.now();

if (datetime.today().weekday() == 6):
  date = datetime.now() - timedelta(days=2)

m = date.strftime("%m")
d = date.strftime("%d")

workbook = excelapp.Workbooks.Open(fr"C:\Users\u0198\Desktop\AutoDev\udn_scrap\0 確診隔離人數統計{m}{d}.xlsx")

sheet1 = workbook.Worksheets(1)
sheet1.Range("B1:H13").Copy()


# In[2]:


from win32com.client import constants 

powerpoint_object = client.Dispatch("Powerpoint.Application")
# powerpoint_object.visible = True

m = datetime.now().strftime("%m")
d = datetime.now().strftime("%d")

powerpoint_presentation = powerpoint_object.Presentations.Open(fr'C:\Users\u0198\Desktop\AutoDev\udn_scrap\新冠即時訊息({m}{d}).pptx')

ptpaste = powerpoint_presentation.slides[3].Shapes.PasteSpecial(DataType=2)
ptpaste.ScaleWidth(1.8, 2.3)
ptpaste.ScaleHeight(1.8, 2.3)


# In[6]:


# from datetime import datetime, timedelta

# date = datetime.now()

# if (datetime.today().weekday() == 2):
#   date = datetime.now() - timedelta(days=2)

# m = date.strftime("%m")
# d = date.strftime("%d")

# m,d

