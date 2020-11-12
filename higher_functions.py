
from lower_functions import *

# ---------------------------------FUNCIÓN: DANIELAPPTX-------------------------------------

def DanielaPPTX(Preferencias):

    #Nombre de la Empresa
    Empresa = {"Nombre":Preferencias[2]}

    #Presentacion base (abrir)
    prs = Presentation('PPTX/' + Empresa["Nombre"].replace('\n','') + '.pptx')
    print('Archivo .pptx cargado')

    #Carga archivo .xlsx
    sheet_ranges = loadXLSX(Preferencias, Empresa)
    print('Archivo .xlsx leído')

    #Encontrar numero de entradas
    index_final = f_FinalIndex(sheet_ranges)

    print('Indice final encontrado: ' + str(index_final-1))

    #Descargar las fotos
    f_loopScreenShots(Empresa["Nombre"], index_final, sheet_ranges)

    print('Fotos descargadas')

    #Comprimir todas las fotos
    f_loopCompress()

    print('Fotos comprimidas')

    index_list = [index_sort[0] for index_sort in sortIndex(index_final, sheet_ranges)]

    for index_copy in index_list:

        #Medidas para los iconos
        IconoMedidas = f_IconoMedidas(Empresa["Nombre"])
        #Medidas para la tabla
        TablaMedidas = f_TablaMedidas(Empresa["Nombre"])
        #Medidas para la captura
        CapturaMedidas = f_CapturaMedidas(Empresa["Nombre"])
        #Medidas para la fecha
        FechaMedidas = f_FechaMedidas(Empresa["Nombre"])

        slide = addSlide(prs)
        #print('Slide añadida')

        table = createTable(slide, TablaMedidas, sheet_ranges, index_copy)
        #print('Tabla añadida')

        addIcon(sheet_ranges['G' + str(index_copy)].value, IconoMedidas, slide)
        #print('Icono añadido')

        addTitle(sheet_ranges['G' + str(index_copy)].value, slide)
        #print('Titulo añadido')

        addDate(sheet_ranges['F' + str(index_copy)].value.strftime("%b %d"),
                FechaMedidas, slide)
        #print('Fecha añadida')

        addLink(sheet_ranges['AA' + str(index_copy)].value, slide)
        #print('Link añadido')

        addCaptura(Empresa["Nombre"], index_copy - 8, slide, CapturaMedidas)
        #print('Captura añadida')

        print('Completado entrada #', str(index_copy))

    prs.save("Presentacion Terminada.pptx")
    os.remove(Preferencias[1] + '.xlsx')

    return os.path.getsize('Presentacion Terminada.pptx')/1000000


# --------------------------------FUNCION: DANIELAWORD---------------------------------------

def DanielaWord(Preferencias):

    # Abrir archivo .docx
    doc = Document(Preferencias[1]+ '.docx')
    print('Documento  cargado: ' + Preferencias[1] + '.docx')

    # Abrir archivo .xlsx base y cargar hoja de calculo
    wb = load_workbook(filename='Excel Base.xlsx')
    sheet_ranges = wb.active

    # Convertir formato de parrafos a texto unico
    txt = prf_2_txt(doc)
    print('Conversion parrafos -> string')

    orderGestion = SortCategories(txt)

    # Conseguir la lista de categorias
    categorias = load_categorias()

    # Conseguir la data
    data = f_extractElement_loop(categorias, txt)

    data = BinaryUpdate_loop(data)

    data = f_date2number_loop(data)

    data = addIndex(data)

    FillExcel(data, categorias, sheet_ranges, orderGestion)
    print('Excel completado')

    wb.save(filename='Excel Terminado.xlsx')
    os.remove(Preferencias[1]+'.docx')

    print('Excel guardado')

    return os.path.getsize('Excel Terminado.xlsx')/1000000

# --------------------------------FUNCION: DOWNLOAD_EMAIL-----------------------------------

def Download_Email():

    mail = f_Email_Preferencias()

    msgs = get_messages(mail)

    if len(msgs) > 0:

        for email_id in msgs:

            Preferencias = get_attachments(mail, msgs, email_id)

        CheckEmail = True

    else:

        CheckEmail = False
        Preferencias = []

    return Preferencias, CheckEmail

# --------------------------------FUNCION: DELETE_EMAIL-----------------------------------

def Delete_Email():

    mail = f_Email_Preferencias()

    msgs = get_messages(mail)

    for num in msgs:
        mail.store(num, '+FLAGS', '\\Deleted')

    mail.expunge()
    mail.close()
    mail.logout()


# --------------------------------FUNCION: SEND_EMAIL-----------------------------------

def Send_Email(tipo, file_size):

    Cred = Credenciales_Correo()

    msg, filename = CreateMessage(file_size, tipo, Cred)

    part = addAttachment(msg, filename)

    txt = encodeMsg(part, msg)

    SecureSend(txt, Cred)
