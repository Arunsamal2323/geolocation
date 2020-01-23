from django.shortcuts import render
from django.http import HttpResponse, Http404
from geopy.geocoders import Nominatim
from tablib import Dataset
from .tests import UploadFileForm
import os.path
# from openpyxl import Workbook
import openpyxl
from io import BytesIO




def excel_upload(request):
    #Path of the file
    
    PROJECT_ROOT = os.path.abspath(os.path.dirname(__file__))
    if request.method == 'POST':
        xfile=UploadFileForm(request.POST,request.FILES)
        #If not valid return the error
        if xfile.is_valid():
            xfile=xfile.cleaned_data['docfile']
        else:
            return HttpResponse(xfile._errors.as_text)
        
        rdata = openpyxl.load_workbook(filename=BytesIO(xfile.read()))
        gdata = rdata.active
        f = gdata['A']
        alldata = [f[i].value for i in range(len(f))]

        geolocator = Nominatim(user_agent="Gmaps")
        addrs = []
        lat = []
        long = []

        for loc in alldata:
            print(loc)
            location = geolocator.geocode(loc, timeout=15)
            if location is not None :
                if location.latitude is not None:
                    addrs.append(loc)
                    lat.append(location.latitude)
                    long.append(location.longitude)
                else:
                    return HttpResponse('Not Available')
        #new file 
        wb2=openpyxl.Workbook()
        sheet = wb2.active
        col= ['Address','Latitude','Longitude']
        sheet.append(col)

        for i in range(len(addrs)):
           sheet.append([addrs[i],lat[i],long[i]])
        #    print([addrs[i],lat[i],long[i]])
        
        wb2.save(os.path.join(PROJECT_ROOT,'output.xlsx'))

        # data.to_excel(os.path.join(PROJECT_ROOT,'output.xlsx'))

    return render(request,'upload.html',context={'form':UploadFileForm()})

def download(request):
    PROJECT_ROOT = os.path.abspath(os.path.dirname(__file__))
    file_path = os.path.join(PROJECT_ROOT,'output.xlsx')
    if os.path.exists(file_path):
        with open(file_path, 'rb') as fh:
            response = HttpResponse(fh.read(), content_type="application/vnd.ms-excel")
            response['Content-Disposition'] = 'inline; filename=' + os.path.basename(file_path)
            return response
    raise Http404
