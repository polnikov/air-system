class CONSTANTS:
    APP_TITLE = 'Расчет естественной вентиляции'
    TAB1_TITLE = "Основной расчёт"
    TAB2_TITLE = "Расчёт дефлектора"
    ACCELERATION_OF_GRAVITY = 9.81

    RESULT_ICONS = (
        './icons/ok.png',
        './icons/no.png',
    )

    MENU = (
        'Файл',
        'Руководство',
        'Помощь',
    )
    FILE_SUBMENU = (
        'Открыть',
        'Сохранить',
        'Сохранить как...',
        'Экспорт в DOCX'
    )
    FILE_MENU_ICONS = (
        './icons/open.png',
        './icons/save.png',
        './icons/save_as.png',
        './icons/export.png',
    )
    FILE_MENU_SHORTCUTS = (
        'Ctrl+O',
        'Ctrl+S',
        'Ctrl+Shift+S',
        'Ctrl+E',
    )
    HELP_SUBMENU = (
        'О программе',
        'Проверить обновления',
    )
    HELP_MENU_ICONS = (
        './icons/about.png',
        './icons/update.png',
    )
    ABOUT = '''<html>
    📄 Программа основана на методике...<hr>
    💡 Методика: Константин @ <font color="blue">nedich@mail.ru</font><br>
    🛠 Программная часть: Юрий @ <font color="blue">akudja.technology@gmail.com</font><hr>
    Почта для предложений и замечаний: @ <font color="blue">akudja.technology@gmail.com</font><hr>
    <center>🤑 Или можете угостить нас 🍺: ХХХХ ХХХХ ХХХХ ХХХХ</center>
    </html>
    '''


    class INIT_DATA:
        TITLE = 'Исходные данные'
        LINE_HEIGHT = 35
        INPUT_WIDTH = 50
        LABELS = (
            ('Расчетная температура внутреннего воздуха, tв', '°C'),
            ('Коэффициент эквивалентной шероховатости\nканала, Kэ', 'мм'),
            ('Высота типового этажа, hэ', 'м'),
            ('Высота шахты, Hш', 'м'),
        )
        KLAPAN_LABEL = 'Воздушный клапан'
        KLAPAN_ITEMS = {
                'Выбрать': '',
                # 1
                'КИВ-125 (КПВ-125)': 36,
                # 3
                'Air-Box Comfort': 31,
                'Air-Box Comfort\nс козырьком': 42,
                'Air-Box Comfort S': 41,
                'Air-Box Comfort Eco': 26,
                # 8
                'Norvind pro': 32,
                'Norvind optima': 13,
                'Norvind classic': 16,
                'Norvind city': 30,
                'Norvind lite': 26,
                # 14
                'Aereco EHT² 5-40': 40,
                'Aereco EHT² 11-40': 40,
                'Aereco EHT² 6-30': 30,
                'Aereco EFT² 24': 24,
                'Aereco EFT² 40': 40,
                'Aereco EFTO² 40': 40,
                # 21
                'Aereco EHT 5-40': 40,
                'Aereco EHT 11-40': 40,
                'Aereco EHT 11-40': 40,
                'Aereco EFT 40': 40,
                # 25
                'Aereco EHA² 5-35': 35,
                'Aereco EHA² 11-35': 35,
                'Aereco EHA² 17-35': 35,
                # 29
                'Aereco EMM² 5-35': 35,
                'Aereco EMM² 11-35': 35,
                'Aereco EMM² 24': 24,
                'Aereco EMM² 35': 35,
                'Aereco EMM² 5-35\nс проставкой': 45,
                'Aereco EMM² 11-35\nс проставкой': 45,
                'Aereco EMM² 24\nс проставкой': 36,
                'Aereco EMM² 35\nс проставкой': 45,
                # 38
                'Aereco EAH² 5-35': 35,
                'Aereco EAH² 11-35': 35,
                'Aereco EAH² 24': 24,
                'Aereco EAH² 35': 35,
                'Aereco EAH² 5-35\nс проставкой': 50,
                'Aereco EAH² 11-35\nс проставкой': 50,
                'Aereco EAH² 24\nс проставкой': 38,
                'Aereco EAH² 35\nс проставкой': 50,
                # 47
                'Aereco EMM 5-35': 35,
                'Aereco EMM 11-35': 35,
                'Aereco EMF 35': 35,
                # 51
                'Другой': '--',
            }
        KLAPAN_INPUT_LABEL_1 = 'Расход воздуха через клапан при перепаде\nдавления 10 Па'
        KLAPAN_INPUT_TOOLTIP = 'Активируется при выборе\nвоздушного клапана <Другой>'
        SURFACE_INPUT_TOOLTIP = '''<html>
            0.1____Сталь, винипласт<br>
            0.11___Асбест<br>
            1______Шлако-алебастр<br>
            1.5____Шлакобетон<br>
            4______Кирпич<br>
            10_____Штукатурка по сетке
        </html>'''


    class BUTTONS:
        TITLE = 'Действия'
        DELETE_BUTTON_TITLE = 'Удалить\nэтаж'
        ADD_BUTTON_TITLE = 'Добавить\nэтаж'
        LAST_FLOOR_BUTTON_TITLE = 'Добавить\nпоследний этаж'
        DELETE_BUTTON_TOOLTIP = 'Удалить текущую строку основного расчёта'
        ADD_BUTTON_TOOLTIP = 'Добавить строку в конец основного расчёта'
        ADD_FLOOR_FOR_DELETE_TOOLTIP = 'Введите этаж, который нужно удалить'
        ADD_DEFLECTOR = 'Добавить дефлектор'


    class CAP:
        NAME = 'cap'
        LABEL_1 = 'Тип оголовка шахты:'
        LINE_HEIGHT = 30
        TYPES = (
            'Выбрать',
            'Без оголовка',
            'Зонт',
            'Плоский экран',
        )
        TYPES_IMG = (
            'icons/zont.png',
            'icons/plate.png',
        )
        RELATIONS = {
            'Зонт': {
                'Выбрать': '',
                'h/Do: ζ': '',
                '0.10: 4.00': 4.00,
                '0.20: 2.30': 2.30,
                '0.25: 1.90': 1.90,
                '0.30: 1.60': 1.60,
                '0.35: 1.40': 1.40,
                '0.40: 1.30': 1.30,
                '0.50: 1.15': 1.15,
                '0.60: 1.10': 1.10,
                '0.80: 1.00': 1.00,
                '1.00: 1.00': 1.00,
            },
            'Плоский экран': {
                'Выбрать': '',
                'h/Do: ζ': '',
                '0.25: 3.40': 3.40,
                '0.30: 2.60': 2.60,
                '0.35: 2.10': 2.10,
                '0.40: 1.70': 1.70,
                '0.50: 1.40': 1.40,
                '0.60: 1.20': 1.20,
                '0.80: 1.10': 1.10,
                '1.00: 1.00': 1.00,
            }
        }
        FACT_RELATION_TOOLTIP = 'Фактическое соотношение'
        PRESSURE_TOOLTIP = 'Потери давления на оголовке шахты'
        INPUT_h_TOOLTIP = '0...2 м'

    class MAIN_TABLE:
        HEADER_HEIGHT = 75
        HEIGHT = 30
        WIDTHS = {
            0: 40,
            1: 70,
            2: 70,
            3: 70,
            4: 60,
            5: 60,
            6: 60,
            7: 60,
            8: 60,
            9: 60,
            10: 110,
            11: 110,
            12: 70,
            13: 70,
            14: 70,
            15: 70,
            16: 70,
            17: 70,
            18: 70,
            19: 70,
            20: 70,
        }
        TITLE = 'Основной расчёт'
        HEADER_NAME = 'header'
        ROW_NAME = 'row'
        LAST_ROW_NAME = 'last_row'
        LABELS = [
            'Этаж',
            'Высота\nэтажа\n[м]',
            'Участок',
            'Lрасч\n[м3/ч]',
            'hрасч\n[м]',
            'Pгр\n[Па]',
            'Рдеф\n[Па]',
            'Ррасп\n[Па]',
            'ζпр',
            'ζотв',
            'Сторона канала\na\n(большая)\n[мм]',
            'Сторона канала\nb\n(меньшая)\n[мм]',
            'v\n[м/с]',
            'Dэкв\n[м]',
            'R\n[Па/м]',
            'm',
            'R∙l∙m\n[Па]',
            'Рд\n[Па]',
            'ΔPпр\n[Па]',
            'ΔPотв\n[Па]',
            'ΔP\n[Па]',
            'Результат',
        ]
        TOOLTIP_H = '''<html><body style="background-color: black"><font color="white">
        Расстояние по вертикали от центра воздухозаборного\nустройства до верха вытяжной шахты
        </font></body></html>'''
        TOOLTIP_h = 'Высота центра вытяжной решетки рассматриваемого этажа'


    class SPUTNIK_TABLE:
        HEADER_HEIGHT = 90
        HEIGHT = 30
        WIDTHS = {
            0: 80,
            1: 70,
            3: 90,
            4: 90,
            'other': 72,
        }
        TITLE = 'Расчёт спутника'
        NAME = 'sputnik'
        LABELS = [
            'Участок',
            'L\n[м3/ч]',
            'Длина\nканала\n[м]',
            'Сторона\nканала\na\n(большая)\n[мм]',
            'Сторона\nканала\nb\n(меньшая)\n[мм]',
            'v\n[м/с]',
            'Dэкв\n[м]',
            'R\n[Па/м]',
            'm',
            'R∙l∙m\n[Па]',
            'Рд\n[Па]',
            '∑ζ',
            'Z\n[Па]',
            'Z+R∙l∙m\n[Па]',
            'Выбор\nрасчёта',
        ]
        RADIO_TOOLTIPS = {
            3: 'Расчёт одностороннего блока',
            5: 'Расчёт двухстороннего блока',
        }
        KLAPAN_FLOW_TOOLTIP = 'Проектный расход\nчерез клапан'
        KLAPAN_LABEL = 'Приточный\nклапан'
        SECTORS = {
            2: '1-2',
            4: '1*-2*',
        }


    class REFERENCE_DATA:
        class M:
            X = [100, 150, 200, 250, 300, 350, 400, 450, 500, 550, 600, 650, 700, 750, 800, 900, 1000, 1500]  # 18
            Y = [100, 200, 300, 400, 500, 600, 700, 800, 900, 1000, 1500]  # 11
            line_1 = [1.13,1.15,1.2,1.25,1.3,1.35,1.4,1.45,1.5,1.55,1.6,1.65,1.7,1.75,1.8,1,1,1]
            line_2 = [1.2,1.145,1.13,1.145,1.16,1.18,1.2,1.23,1.25,1.27,1.3,1.33,1.36,1.39,1.41,1.46,1.51,1.75]
            line_3 = [1.3,1.2,1.16,1.14,1.13,1.14,1.148,1.165,1.175,1.195,1.21,1.22,1.245,1.26,1.275,1.32,1.355,1.54]
            line_4 = [1.4,1.265,1.2,1.17,1.148,1.137,1.13,1.137,1.145,1.149,1.16,1.165,1.17,1.185,1.2,1.225,1.25,1.38]
            line_5 = [1.5,1.345,1.25,1.205,1.175,1.1505,1.145,1.135,1.13,1.135,1.14,1.145,1.15,1.155,1.165,1.18,1.2,1.3]
            line_6 = [1.6,1.405,1.3,1.25,1.21,1.18,1.16,1.148,1.14,1.135,1.13,1.135,1.138,1.144,1.148,1.155,1.17,1.25]
            line_7 = [1.7,1.48,1.36,1.29,1.245,1.21,1.17,1.16,1.15,1.145,1.138,1.135,1.13,1.132,1.138,1.145,1.15,1.22]
            line_8 = [1.8,1.55,1.41,1.33,1.275,1.24,1.2,1.18,1.165,1.15,1.148,1.141,1.138,1.138,1.13,1.138,1.146,1.2]
            line_9 = [1,1.6,1.46,1.38,1.32,1.27,1.225,1.2,1.18,1.17,1.155,1.149,1.145,1.147,1.138,1.13,1.137,1.17]
            line_10 = [1,1.68,1.51,1.42,1.355,1.3,1.25,1.225,1.2,1.18,1.17,1.16,1.15,1.146,1.146,1.137,1.13,1.155]
            line_11 = [1,1,1.75,1.62,1.54,1.46,1.38,1.345,1.3,1.275,1.25,1.23,1.22,1.2,1.2,1.17,1.155,1.13]
            TABLE = [line_1, line_2, line_3, line_4, line_5, line_6, line_7, line_8, line_9, line_10, line_11]


        class DEFLECTOR_PRESSURE_RELATION:
            X = [0, 0.05, 0.1, 0.15, 0.2, 0.25, 0.3, 0.35, 0.4, 0.45, 0.5]
            TABLE = [0.55, 0.48, 0.43, 0.38, 0.35, 0.32, 0.28, 0.24, 0.21, 0.16, 0.1]


    class DEFLECTOR:
        TITLE = 'Расчёт дефлектора'
        NAME = 'deflector'
        TOOLTIP = 'Активируется, если расчет\nдефлектора выполнен'
        LABELS = (
            'Скорость ветра Vв [м/с]',
            'Рекомендуемая скорость в патрубке дефлектора Vд.рек [м/с]',
            'Расход воздуха L [м3/ч]',
            'Требуемая площадь патрубка дефлектора Fрек [м2]',
            'Диаметр дефлектора [мм]',
            'Фактическая скорость в патрубке дефлектора Vд.факт [м/с]',
            'Отношение Vд.факт / Vв',
            'Отношение Pд / Pв',
            'Разрежение в патрубке дефлектора Pд [Па]',
        )
        DIAMETERS = {
            0.007854: '100',
            0.012272: '125',
            0.020106: '160',
            0.031416: '200',
            0.049087: '250',
            0.077931: '315',
            0.125664: '400',
            0.19635: '500',
            0.311725: '710',
            0.395919: '800',
        }
