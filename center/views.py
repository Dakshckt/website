from django.shortcuts import render
from django.http import HttpResponse , FileResponse
from datetime import datetime
from docx import Document
from docx2pdf import convert
from .file import replace_table , replace_text
import os

# Create your views here.
def home(req):
    numbers = [1,2,3,4,5]
    sum = 0
    ftype = ""

    if req.method == "POST":

        BASE_DIR = os.path.dirname(os.path.abspath(__file__))  
        PROJECT_ROOT = os.path.dirname(BASE_DIR) 
        input_file = os.path.join(PROJECT_ROOT, "source", "base.docx")
        doc = Document(input_file)
        
        file_type = int(req.POST.get("type"))
        if file_type == 0:
            ftype = "Bill"
        else:
            ftype = "Quatation"
        replace_text(doc , "type" , ftype)
        building = req.POST.get("building")
        replace_text(doc , "building" , building.title())
        city = req.POST.get("city")
        replace_text(doc , "city" , city.title())
        subject = req.POST.get("subject")
        replace_text(doc , "subject" , subject.title())
        curr_date = datetime.now().date()
        replace_text(doc , "date" , str(curr_date))

        for i in numbers:
            rd = req.POST.get(f"rd{i}")
            rr = req.POST.get(f"rr{i}")
            ra = req.POST.get(f"ra{i}")
            rt = req.POST.get(f"rt{i}")
            if rr != "":  
                rr = f"{rr} /-"
            if rt != "":
                sum = sum + float(rt)
                rt = f"{rt} /-"

            replace_table(doc , f"rd{i}" , rd.title())
            replace_table(doc , f"rr{i}" , rr)
            replace_table(doc , f"ra{i}" , ra)
            replace_table(doc , f"rt{i}" , rt)

        if file_type == 0:
            replace_text(doc , "total" , sum)
        else:
            replace_text(doc , "total" , "")

        output_docx = os.path.join(PROJECT_ROOT, f"{ftype}.docx")

        doc.save(output_docx)  

        return FileResponse(open(output_docx, 'rb'), as_attachment=True, filename=f"{building}{ftype}{curr_date}.docx")

    return render(req , "index.html" , {"numbers":numbers})


