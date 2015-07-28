# -*- coding: utf-8 -*-
import io
import mimetypes
import os
import tempfile
import uuid
from wsgiref.util import FileWrapper
import zipfile
from django.http import HttpResponse, StreamingHttpResponse
from django.views.decorators.http import require_POST, require_GET
import openpyxl
from openpyxl.worksheet import Selection
import urllib.parse
import before_nohin.zipfile_cp932

__author__ = 'koty'
from django.shortcuts import render

xlsx_temp_filename = "xlsfile.xlsx"
zip_temp_filename = "xlsfile.zip"

@require_GET
def index(request):
    return render(request, 'index.html', {})


@require_POST
def upload(request):
    f = request.FILES.get('xlsfile')
    if not f:
        return render(request, 'index.html', {'message': 'ファイルを指定してください。'})

    temp_dir = tempfile.gettempdir() + '/before_nohin/' + uuid.uuid4().hex + '/'
    os.makedirs(temp_dir)
    if f.name.lower().endswith('.zip'):
        if not _write_zipfile(f, temp_dir):
            return render(request, 'index.html', {'message': '不正なzipファイルです。'})
        zip_path = _handle_extracted_files(temp_dir, f.name)
        return _get_zip(zip_path)
    else:
        filename = _write_tempfile(f, temp_dir)
        output = io.BytesIO()
        _handle_xls(output, filename)
        return _get_xlsx(output, f.name)


def _write_tempfile(f, temp_dir):
    filename = temp_dir + xlsx_temp_filename
    with open(filename, 'wb+') as destination:
        os.makedirs(temp_dir + 'source/')
        for chunk in f.chunks():
            destination.write(chunk)
    return filename

def _write_zipfile(f, temp_dir):
    zip_full_path = temp_dir + "source/" + zip_temp_filename
    os.makedirs(temp_dir + 'source/')
    with open(zip_full_path, 'wb+') as destination:
        for chunk in f.chunks():
            destination.write(chunk)
    if not zipfile.is_zipfile(zip_full_path):
        return False

    with zipfile.ZipFile(zip_full_path, 'r') as zf:
        # protect extraction to unexpected location
        invalid_filenames = [zfile.filename for zfile in zf.filelist
                             if '..' in zfile.filename or zfile.filename.startswith('/')]
        if invalid_filenames:
            return False
        zf.extractall(temp_dir + "source/")

    return True

def _handle_extracted_files(temp_dir, zip_filename):
    source_path = temp_dir + "source/"
    os.makedirs(temp_dir + 'target/')
    for root, subdirs, files in os.walk(source_path):
        for filename in [f for f in files if f.endswith('.xlsx')]:
            xls_full_path = os.path.join(root, filename)
            wb = openpyxl.load_workbook(filename=xls_full_path)

            for worksheet in wb.worksheets:
                _reset_sheet(worksheet)

            wb.active = 0

            target_root_path = root.replace('/source/', '/target/')
            if not os.path.exists(target_root_path):
                os.makedirs(target_root_path)
            xls_target_path = os.path.join(target_root_path, filename)
            wb.save(xls_target_path)

    zip_target = temp_dir + zip_filename
    with zipfile.ZipFile(zip_target, 'w') as zip_file:
        zipdir(temp_dir + "target/", zip_file)

    return zip_target

def zipdir(path, zip_file):
    for root, dirs, files in os.walk(path):
        for file in files:
            zip_file.write(os.path.join(root, file), arcname=file.replace(path, ''))


def _handle_xls(output, file_name):
    # ファイルの読み込み
    wb = openpyxl.load_workbook(filename=file_name)

    for worksheet in wb.worksheets:
        _reset_sheet(worksheet)

    wb.active = 0
    wb.save(output)

def _get_xlsx(output, filename):
    response = HttpResponse(output.getvalue(),
                            content_type=mimetypes.guess_type(filename)[0])
    response['Content-Disposition'] = "attachment; filename*=UTF-8''{0}".format(urllib.parse.quote(filename))
    return response

def _get_zip(filename):
    basename = os.path.basename(filename)
    o = open(filename, mode='br')
    file_wrapper = FileWrapper(o)
    response = StreamingHttpResponse(file_wrapper,
                                     content_type=mimetypes.guess_type(filename)[0])
    response['Content-Length'] = os.path.getsize(filename)
    response['Content-Disposition'] = "attachment; filename*=UTF-8''{0}".format(urllib.parse.quote(basename))

    return response

def _reset_sheet(worksheet):
    """
    set zoom 100% and selected cell A1
    :param worksheet:
    :return:
    """
    view = worksheet.sheet_view
    view.zoomScale = "100"
    view.zoomScaleNormal = "100"
    view.zoomScalePageLayoutView = "100"
    view.selection = (Selection(activeCell="A1", sqref="A1"),)
