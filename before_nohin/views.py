# -*- coding: utf-8 -*-
import io
from django.http import HttpResponse
from django.views.decorators.http import require_POST, require_GET
import openpyxl

__author__ = 'koty'
from django.shortcuts import render

filename = "/tmp/xlsfile.xlsx"

@require_GET
def index(request):
    return render(request, 'index.html', {})


@require_POST
def upload(request):
    f = request.FILES['xlsfile']
    _write_tempfile(f)
    output = io.BytesIO()
    _handle_xls(output)
    return _get_xlsx(output)


def _write_tempfile(f):
    destination = open(filename, 'wb+')
    for chunk in f.chunks():
        destination.write(chunk)
    destination.close()

def _handle_xls(output):
    # テンプレートファイルの読み込み
    wb = openpyxl.load_workbook(filename=filename)

    wb.worksheets[0].active = True
    sheet = wb.worksheets[0]
    sheet["A2"].value = 10
    sheet.range('A1')

    wb.save(output)

def _get_xlsx(output):
    response = HttpResponse(output.getvalue(), content_type="application/excel")
    response['Content-Disposition'] = "filename=xlsfile.xlsx"
    return response
