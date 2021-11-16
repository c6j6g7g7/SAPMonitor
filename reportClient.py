import datetime

from reportlab.platypus import SimpleDocTemplate
from reportlab.lib.pagesizes import letter, B6, landscape

from reportlab.platypus import Table
from reportlab.platypus import Image
from reportlab.platypus import TableStyle
from reportlab.lib import colors


from pdf2image import convert_from_bytes

class report_table:
    title = "Reporte inicial de servicios"
    date = datetime.datetime.now()

class ReportClient:



    def __init__(self, system_data, client, path_gen):
        self.system_data = system_data
        self.fileName = path_gen+client+'.pdf'
        self.path_gen = path_gen
        self.client = client
        self.data_table = []
        self.gray = colors.Color(red=(191 / 255), green=(191 / 255), blue=(191 / 255))
        self.green = colors.Color(red=(169 / 255), green=(208 / 255), blue=(142 / 255))
        self.yellow = colors.Color(red=(255 / 255), green=(192 / 255), blue=(0))  # blue=(0/255))
        self.red = colors.Color(red=(255 / 255), green=(0), blue=(0))  # green=(0/255), blue=(0))


    def header_report(self,title):
        elemWidth = 250

        # 1 - Build Structure
        titleTable = Table([
            [title]
        ], elemWidth)

        date = datetime.datetime.now()
        dateTable = Table([
            [ date.strftime('%d/%m/%Y') ]
        ], [60])

        elemTable = Table([
            [titleTable, dateTable]
        ], elemWidth)

        # 2 - Build Style
        style = TableStyle([
            ('FONTSIZE', (0, 0), (0, 0), 12),
            ('FONTNAME', (0, 0), (0, 0),'Helvetica-Bold'),
            ('ALIGN', (0, 0), (0, 0), 'LEFT'),
            ('TOPPADDING', (0, 0), (-1, -1), 0),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 0),

        ])
        titleTable.setStyle(style)

        dateTable.setStyle(style)

        return elemTable

    def footer_report(self):
        pinElemTable = None
        # conventionsTable

        titleTable = Table([
            ['Convenciones']
        ], 210)

        conventionsTable = Table([
            ['Operativo', 'Cambio', 'No Operativo']
        ], [70, 70, 70])

        # Style
        # Para cliente
        style = TableStyle([
            ('BACKGROUND', (0, 0), (0, 0), self.gray),
            ('FONTSIZE', (0, 0), (0, 0), 8),
            ('FONTNAME', (0, 0), (0, 0), 'Helvetica-Bold'),
            ('ALIGN', (0, 0), (0, 0), 'LEFT'),
            ('TOPPADDING', (0, 0), (-1, -1), 0),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 0),

        ])
        titleTable.setStyle(style)

        style = TableStyle([
            ('FONTSIZE', (0, 0), (-1, -1), 8),
            ('FONTNAME', (0, 0), (-1, -1), 'Helvetica-Bold'),
            ('TEXTCOLOR', (0, 0), (-1, -1), colors.whitesmoke),
            ('BACKGROUND', (0, 0), (0, 0), self.green),
            ('BACKGROUND', (1, 0), (1, 0), self.yellow),
            ('BACKGROUND', (2, 0), (2, 0), self.red),
            ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ])

        conventionsTable.setStyle(style)

        pinElemTableStyle = TableStyle([
            #('BOX', (0, 0), (-1, -1), 1, colors.black),
            ('TOPPADDING', (0, 0), (-1, -1), 0),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 0),
            ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ])

        ElemTable = Table([
            [titleTable]
            , [conventionsTable]
        ], 210)

        ElemTable.setStyle(pinElemTableStyle)

        return ElemTable

    def gen_enviroment_header(self):
        data = ['']
        for system in self.system_data:
            if (system['enviroment'] not in data):
                data.append(system['enviroment'])
        return data

    def gen_systems_header(self, client):
        data = ['','Instance']
        tmp = []
        for system in self.system_data:
            if (system['system'] not in data):
                data.append(system['system'])
        tmp.append(client)
        for i in range(len(data) - 1):
            tmp.append('')
        self.data_table.append(tmp)
        self.data_table.append(data)

    def gen_systemids_header(self):
        data_env = self.gen_enviroment_header()
        last_enviroment = None
        xi = 0
        yi = 2
        posx = 1
        y = 2
        sum = 0
        style = []
        aling_style = []
        for env in data_env[1:]:
            env_tmp = env
            y = 2
            sum = 0
            for system in self.system_data:
                if env in system['enviroment']:
                    lista = []
                    for i in range(len(self.data_table[1])):
                        lista.append('')

                    if env_tmp in self.data_table:
                        lista[0] = ''
                        lista[1] = ''
                    else:
                        lista[0] = system['enviroment']
                        lista[1] = system['instance']

                    pos = self.data_table[1].index(system['system'])
                    lista[pos] = system['systemid']

                    posx = lista.index(system['systemid'])
                    y = len(self.data_table)

                    if system['connection'] == 'OK':
                        tupla_style = ('BACKGROUND', (posx, y), (posx, y), self.green)
                    else:
                        tupla_style = ('BACKGROUND', (posx, y), (posx, y), self.red)

                    style.append(tupla_style)

                    self.data_table.append(lista)

        return TableStyle(style), self.gen_style_enviroments_data()

    def gen_style_enviroments_data(self):
        xi = 0
        xf = 0
        yi = 0
        yf = 0
        x = 0
        y = 0
        last_env = None
        env_list = {}
        list_style = []
        align_style = []
        for y, env in enumerate(self.data_table[1:]):
            if y == 0:
                anterior = env
            for x, data in enumerate(env):
                if x == 0:
                    if anterior[x] in env[x]:
                        xi = x
                        yi = y

                    if anterior[x] != env[x]:
                        xf = x
                        yf = y
                        tupla = ['SPAN', (xi, yi), (xf, yf)]
                        tupla_alig = ('VALIGN', (xi, yi), (xf, yf), 'MIDDLE')
                        list_style.append(tupla)
                        align_style.append(tupla_alig)
                        xi = 0
                        yi = y + 1
            anterior = env
        return TableStyle(list_style)


    def gen_style_systemids_data(self):
        xi = 1
        xf = 0
        yi = 2
        yf = 0
        x = 0
        y = 0
        last_env = None
        env_list = {}
        list_style = []
        align_style = []
        for y, env in enumerate(self.data_table[2:]): ######
            if y == 0:
                anterior = env
            print("ENV-> {}".format(env))
            yi = None
            for x, data in enumerate(env):
                if x == 1: ##########
                    if anterior[x] in env[x] and yi == None:
                        print("ESTA EL ANTERIOR->{}".format(x))
                        xi = x
                        yi = y + 2 #######

                    if anterior[x] != env[x]:
                        print("anterior[x] ->{} != env[x]->{}".format(anterior[x], env[x]))
                        xf = x
                        yf = y
                        tupla = ['SPAN', (xi, yi), (xf, yf)]
                        tupla_alig = ('VALIGN', (xi, yi), (xf, yf), 'MIDDLE')
                        #list_style.append(tupla)
                        #align_style.append(tupla_alig)
                        xi = 1 ########
                        #yi = y + 2  ############
                        print("tupla-> {}".format(tupla))
            anterior = env

        tupla = ['SPAN', (xi, yi), (xf, yf)]
        #
        tupla_alig = ('VALIGN', (xi, yi), (xf, yf), 'MIDDLE')
        #print("tupla-> {}".format(tupla))
        list_style.append(tupla)

        return TableStyle(list_style)


    def gen_style_enviroment(self):
        last_enviroment = None
        index = 1
        xi = 0
        yi = 2
        list_style = []
        align_style = []
        for list_table in self.data_table[2:]:
            # Style enviroment
            if list_table[0] != last_enviroment and last_enviroment != None:
                tupla = ['SPAN', (xi, yi), (xi, index)]
                tupla_alig = ('VALIGN', (xi, yi), (xi, index), 'MIDDLE')
                list_style.append(tupla)
                align_style.append(tupla_alig)
                yi = index + 1
            last_enviroment = list_table[0]
            index = index + 1

        tupla = ['SPAN', (xi, yi), (xi, index)]
        tupla_alig = ('VALIGN', (xi, yi), (xi, index), 'MIDDLE')
        list_style.append(tupla)
        align_style.append(tupla_alig)

        return TableStyle(list_style), TableStyle(align_style)

    def data_report(self, client):
        pdf = SimpleDocTemplate(
            self.fileName,
            pagesize=letter
        )

        self.gen_systems_header(client)
        style, style_align = self.gen_systemids_header()

        # Se genera la cabecera del reporte
        header = self.header_report('Report Inicial de Servicios')
        footer = self.footer_report()

        #print("data_table->{}".format(self.data_table))

        table = Table(
            self.data_table
        )

        table.setStyle(style)
        table.setStyle(style_align)

        # Para cliente
        style = TableStyle([
            ('BACKGROUND', (0, 0), (0, 0), self.gray),
            ('FONTNAME', (0, 0), (0, 0),
             'Helvetica-Bold'),
        ])
        table.setStyle(style)

        # Para headears
        style = TableStyle([
            ('BACKGROUND', (1, 1), (-1, 1), self.gray),
            ('TEXTCOLOR', (1, 1), (-1, 1), colors.whitesmoke),
            ('ALIGN', (1, 1), (-1, 1), 'CENTER'),
            ('ALIGN', (2, 1), (-1, -1), 'CENTER'),
        ])
        table.setStyle(style)

        # Para bordes tabla
        style = TableStyle(
            [
                #('GRID', (0, 1), (-1, -1), 2, colors.white),
                ('GRID', (0, 1), (-1, -1), 1, colors.black),
            ]
        )
        table.setStyle(style)

        # Style Enviroment
        style, align = self.gen_style_enviroment()

        table.setStyle(style)
        table.setStyle(align)

        ################PARA LAS pRUEBAS
        #table.setStyle(self.gen_style_systemids_data())
        ###############################

        maintable = Table([
            [header],
            [table],
            [footer]
        ]
        )

        elems = []
        elems.append(maintable)

        pdf.build(elems)

    def gen_image(self, client):
        # Generar imagen apartir de PDF
        #image = convert_from_bytes(open(self.fileName, 'rb').read())[0]
        #image.save(self.path_gen+client + '.png')
        return client + '.png'

    def build_report(self):
        self.data_report(self.client)
        return self.gen_image(self.client)