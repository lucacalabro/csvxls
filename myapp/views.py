import csv
import xlwt
from django.http import HttpResponse
from .models import User
from django.shortcuts import render


def index(request):
    return render(request, "index.html")


# Export dati in formato csv
def export_users_csv(request):
    response = HttpResponse(content_type='text/csv')
    response['Content-Disposition'] = 'attachment; filename="users.csv"'

    writer = csv.writer(response, delimiter=';')
    writer.writerow(['Username', 'First name', 'Last name', 'Email address'])

    users = User.objects.all().values_list('username', 'first_name', 'last_name', 'email')
    for user in users:
        writer.writerow(user)

    return response


# Export dati in formato xls
def export_users_xls(request):
    response = HttpResponse(content_type='application/ms-excel')
    response['Content-Disposition'] = 'attachment; filename="users.xls"'

    wb = xlwt.Workbook(encoding='utf-8')
    ws = wb.add_sheet('Users')

    # Indice dell'header
    row_num = 0

    style = xlwt.XFStyle()
    style.font.bold = True

    columns = ['Username', 'First name', 'Last name', 'Email address', ]

    #Scrittura dell'header
    for col_num in range(len(columns)):
        ws.write(row_num, col_num, columns[col_num], style)

    # Righe rimanenti
    # contenenti i dati estratti dalla tabella
    style = xlwt.XFStyle()

    #Imposto i bordi
    borders= xlwt.Borders()
    borders.left= 0
    borders.right= 0
    borders.top= 2
    borders.bottom= 2


    style.borders = borders



    rows = User.objects.all().values_list('username', 'first_name', 'last_name', 'email')
    for row in rows:
        row_num += 1
        for col_num in range(len(row)):
            ws.write(row_num, col_num, row[col_num], style)

    wb.save(response)
    return response
