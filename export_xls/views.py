import django
from django.http import HttpResponse
import xlwt
from datetime import datetime, date
from django.template.defaultfilters import slugify
from django.conf import settings

#https://www.codeleading.com/article/49641304939/#:~:text=python%20%E6%AF%94%E8%BE%83%E4%B8%A4%E4%B8%AA%E7%89%88%E6%9C%AC%E5%8F%B7%E5%A4%A7%E5%B0%8F%20%E6%8A%80%E6%9C%AF%E6%A0%87%E7%AD%BE%EF%BC%9A%20Python,%E6%AF%94%E8%BE%83%E4%B8%A4%E4%B8%AA%E7%89%88%E6%9C%AC%E5%8F%B7ver1%E5%92%8Cver2%E7%9A%84%E5%A4%A7%E5%B0%8F%201%E3%80%81%E9%A6%96%E5%85%88%E5%B0%86%E4%B8%A4%E4%B8%AA%E7%89%88%E6%9C%AC%E5%8F%B7%E5%A4%84%E7%90%86%E6%88%90%E7%BA%AF%E6%95%B0%E5%AD%97%E7%9A%84%E7%89%88%E6%9C%AC%E5%8F%B7%EF%BC%8C%E5%A6%825.2.1.3.20%202%E3%80%81%E5%B0%86%E7%89%88%E6%9C%AC%E5%8F%B7%E6%8C%89%E2%80%9C.%E2%80%9D%E5%88%87%E5%89%B2%E4%B8%BA%E5%88%97%E8%A1%A8%EF%BC%8C%E4%BB%8E%E7%B4%A2%E5%BC%950%E5%BC%80%E5%A7%8B%E4%BE%9D%E6%AC%A1%E6%AF%94%E8%BE%83%E5%88%97%E8%A1%A8%E7%9A%84%E5%A4%A7%E5%B0%8F%203%E3%80%81%E5%AF%B9%E6%AF%94%E4%B8%A4%E4%B8%AA%E5%88%97%E8%A1%A8%E7%9A%84len%EF%BC%8Clen%E8%BE%83%E7%9F%AD%E7%9A%84%E4%BD%9C%E4%B8%BA%E5%BE%AA%E7%8E%AF%E6%AC%A1%E6%95%B0%EF%BC%8C%E9%98%B2%E6%AD%A2%E5%88%97%E8%A1%A8%E7%B4%A2%E5%BC%95%E8%B6%8A%E7%95%8C%204%E3%80%81%E5%A6%82%E6%9E%9C%E5%BE%AA%E7%8E%AF%E7%BB%93%E6%9D%9F%E5%90%8E%E4%BB%8D%E6%B2%A1%E6%9C%89%E5%AF%B9%E6%AF%94%E5%87%BA%E7%BB%93%E6%9E%9C%EF%BC%8C%E5%88%99%E5%AF%B9%E6%AF%94%E5%88%97%E8%A1%A8len%EF%BC%8Clen%E5%80%BC%E5%A4%A7%E7%9A%84%E4%B8%BA%E9%AB%98%E7%89%88%E6%9C%AC
def compared_version(ver1, ver2):
    """
    传入不带英文的版本号,特殊情况："10.12.2.6.5">"10.12.2.6"
    :param ver1: 版本号1
    :param ver2: 版本号2
    :return: ver1< = >ver2返回-1/0/1
    """
    list1 = str(ver1).split(".")
    list2 = str(ver2).split(".")
    # print(list1)
    # print(list2)
    # 循环次数为短的列表的len
    for i in range(len(list1)) if len(list1) < len(list2) else range(len(list2)):
        if int(list1[i]) == int(list2[i]):
            pass
        elif int(list1[i]) < int(list2[i]):
            return -1
        else:
            return 1
    # 循环结束，哪个列表长哪个版本号高
    if len(list1) == len(list2):
        return 0
    elif len(list1) < len(list2):
        return -1
    else:
        return 1


def export_xlwt(filename, fields, values_list, save=False, folder=""):
    """export_xlwt is a function based on http://reliablybroken.com/b/2009/09/outputting-excel-with-django/"""
    filename = slugify(filename)
    book = xlwt.Workbook(encoding='utf8')
    sheet = book.add_sheet(filename)

    default_style = xlwt.Style.default_style
    datetime_style = xlwt.easyxf(num_format_str='dd/mm/yyyy hh:mm')
    date_style = xlwt.easyxf(num_format_str='dd/mm/yyyy')

    for j, f in enumerate(fields):
        sheet.write(0, j, fields[j])

    for row, rowdata in enumerate(values_list):
        for col, val in enumerate(rowdata):
            if isinstance(val, datetime):
                style = datetime_style
            elif isinstance(val, date):
                style = date_style
            else:
                style = default_style

            sheet.write(row + 1, col, val, style=style)

    if not save:
        dv = django.get_version()
        if compared_version(dv, '1.7')>0:
            response = HttpResponse(content_type='application/vnd.ms-excel')
        else:
            response = HttpResponse(mimetype='application/vnd.openxmlformats-officed')
        response['Content-Disposition'] = 'attachment; filename=%s.xls' % filename
        book.save(response)
        return response
    else:
        dirpath = '%s/%s' % (settings.MEDIA_ROOT, folder)
        if folder != "":
            import os
            if not os.path.exists(dirpath):
                os.makedirs(dirpath)
        filepath = '%s%s.xls' % (dirpath, filename)
        book.save(filepath)
        return HttpResponse("%s%s%s.xls" % (settings.MEDIA_URL, folder, filename))
