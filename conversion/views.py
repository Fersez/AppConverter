__author__ = 'ccialproinco2'
from django.shortcuts import render_to_response, RequestContext
from django.http import HttpResponseRedirect
from django.core.files.storage import default_storage
from django.core.files.base import ContentFile
from django.conf import settings
import os
import xlrd
import xlwt
import time

def index(request):
    args = {}
    error = 0
    if request.method == 'POST':
        file = request.FILES['excel']
        name=file.name
        index = name.rfind('.')+1
        extention = name[index:]
        if extention != 'csv' and extention != 'xls' and extention != 'xlsx' and extention != 'xlsm':
            error = 1
        if error == 0:
            path = default_storage.save('tmp/'+name, ContentFile(file.read()))
            os.path.join(settings.MEDIA_ROOT, path)
            return HttpResponseRedirect('/seleccion/'+name[:index-1]+'/'+extention)

    args['error']=error

    return render_to_response('index.html',args,context_instance=RequestContext(request))

def selection(request, name,extention):
    book = xlrd.open_workbook('tmp/'+name+'.'+extention)
    sh = book.sheet_by_index(0)
    first_line = []
    for i in range(sh.ncols):
        first_line.append(sh.cell_value(rowx=0, colx=i))
    args = {'headers':first_line}
    if request.method == 'POST':
        comunes = request.POST.getlist('comunes')
        atributos = request.POST.getlist('atributos')
        id_atributo = int(request.POST['atributo']) + 1
        caracteristicas = request.POST.getlist('caracteristicas')
        id_caracteristica = int(request.POST['caracteristica']) + 1

        dict_atributos = {}
        dict_caracteristicas = {}
        dict_car_values = {}
        dict_atr_values = {}


        excel_atributos = xlwt.Workbook()
        sheet_atr = excel_atributos.add_sheet('Hoja 1')
        for index, value in enumerate(atributos):
            row = sheet_atr.row(index)
            row.write(0, id_atributo)
            row.write(1, value)
            dict_atributos[value] = id_atributo
            dict_atr_values[id_atributo] = {}
            id_atributo += 1
        excel_atributos.save("atributos_"+time.strftime("%d_%m_%Y_%H_%M_%S")+".xls")

        excel_caracteristicas = xlwt.Workbook()
        sheet_car = excel_caracteristicas.add_sheet("Hoja 1")
        for index, value in enumerate(caracteristicas):
            row = sheet_car.row(index)
            row.write(0, id_caracteristica)
            row.write(1, value)
            dict_caracteristicas[value] = id_caracteristica
            dict_car_values[id_caracteristica] = {}
            id_caracteristica += 1
        excel_caracteristicas.save("caracteristicas_"+time.strftime("%d_%m_%Y_%H_%M_%S")+".xls")


        excel_articulos = xlwt.Workbook()
        sheet_articulos = excel_articulos.add_sheet("Hoja 1")

        excel_atr_valor = xlwt.Workbook()
        sheet_atr_valor = excel_atr_valor.add_sheet("Hoja 1")

        excel_car_valor = xlwt.Workbook()
        sheet_car_valor = excel_car_valor.add_sheet("Hoja 1")

        column_atr = 0
        column_car = 0

        row_atr_value = 1
        row_car_value = 1
        
        for i in range(sh.nrows):
            if i == 0:
                
                row_at = sheet_atr_valor.row(i)
                row_at.write(0, 'id_atributo')
                row_at.write(1, 'id_valor')
                row_at.write(2, 'nombre_valor')
                
                row_car = sheet_car_valor.row(i)
                row_car.write(0, 'id_caracteristica')
                row_car.write(1, 'id_valor')
                row_car.write(2, 'nombre_valor')

                row_art = sheet_articulos.row(i)
                for index, value in enumerate(comunes):
                    row_art.write(index, value)
                length = len(comunes)
                column_atr = length
                row_art.write(column_atr, 'atributos')
                column_car = column_atr + 1
                row_art.write(column_car, 'caracteristicas')
            else:
                car_value_codif = ""
                atr_value_codif = ""

                row_art = sheet_articulos.row(i)
                ncolumn = 0
                for index, value in enumerate(first_line):
                    if value in comunes:

                        row_art.write(ncolumn, sh.cell_value(rowx=i, colx=index))
                        ncolumn += 1


                    if value in atributos:
                        id_atr = dict_atributos[value]
                        atr_value_name = sh.cell_value(rowx=i, colx=index)
                        if atr_value_name != '':
                            if atr_value_name in dict_atr_values[id_atr].keys():
                                id_value = dict_atr_values[id_atr][atr_value_name]
                            else:
                                id_value = len(dict_atr_values[id_atr].keys()) + 1
                                dict_atr_values[id_atr][atr_value_name] = id_value
                                row_at = sheet_atr_valor.row(row_atr_value)
                                row_atr_value += 1
                                row_at.write(0, id_atr)
                                row_at.write(1, id_value)
                                row_at.write(2, atr_value_name)

                            atr_value_codif += str(id_atr)+":"+str(id_value)+";"


                    if value in caracteristicas:
                        id_car = dict_caracteristicas[value]
                        car_value_name = sh.cell_value(rowx=i, colx=index)
                        if car_value_name != '':
                            if car_value_name in dict_car_values[id_car].keys():
                                id_value = dict_car_values[id_car][car_value_name]
                            else:
                                id_value = len(dict_car_values[id_car].keys()) + 1
                                dict_car_values[id_car][car_value_name] = id_value
                                row_car = sheet_car_valor.row(row_car_value)
                                row_car_value += 1
                                row_car.write(0, id_car)
                                row_car.write(1, id_value)
                                row_car.write(2, car_value_name)

                            car_value_codif += str(id_car)+":"+str(id_value)+";"

                row_art.write(column_atr, atr_value_codif)
                row_art.write(column_car, car_value_codif)



        excel_articulos.save("articulos_"+time.strftime("%d_%m_%Y_%H_%M_%S")+".xls")
        excel_atr_valor.save("atributos_valor_"+time.strftime("%d_%m_%Y_%H_%M_%S")+".xls")
        excel_car_valor.save("caracteristicas_valor_"+time.strftime("%d_%m_%Y_%H_%M_%S")+".xls")



    return render_to_response('campos.html', args, context_instance=RequestContext(request))
