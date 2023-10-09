standarts = [
    'ГОСТ 20.57.406-81 «Изделия электронной техники, квантовой электроники и электротехнические. Методы испытаний»;',
    'ГОСТ 2990-78 «Кабели, провода и шнуры. Методы испытания напряжением»;',
    'ГОСТ 3345-76 «Кабели, провода и шнуры. Метод определения электрического сопротивления изоляции»;',
    'ГОСТ 7229-76 «Кабели, провода и шнуры. Метод определения электрического сопротивления токопроводящих жил и проводов»;',
    'ГОСТ 12177-79 «Кабели, провода и шнуры. Методы проверки конструкции»;',
    'ГОСТ 12182.5-80 «Кабели, провода и шнуры. Метод проверки стойкости к растяжению»;',
    'ГОСТ 18690-2012 «Кабели, провода, шнуры и кабельная арматура. Маркировка, упаковка, транспортирование и хранение»;',
    'ГОСТ 27893-88 «Кабели связи. Методы испытаний»;',
    'ГОСТ IEC 60332-3-22-2011 «Испытание электрических и оптических кабелей в условиях воздействия пламени. Часть 3-22. Распространение пламени по вертикально расположенным пучкам проводов или кабелей. Категория А»;',
    'ГОСТ IEC 60811-401-2015 «Кабели электрические и волоконно-оптические. Методы испытаний неметаллических материалов. Часть 401. Раз-ные испытания. Методы теплового старения. Старение в термостате»;',
    'ГОСТ IEC 60811-501-2015 «Кабели электрические и волоконно-оптические. Методы испытаний неметаллических материалов. Часть 501. Ме-ханические испытания. Испытания для определения механических свойств композиций изоляции и оболочки»;',
    'ГОСТ IEC 61034-2-2011 «Измерение плотности дыма при горении кабелей в заданных условиях. Часть 2. Метод испытания и требования к нему».',
]


list_head_SI_IO = [
    'Наименование ИО, СИ',
    'Тип ИО СИ',
    'Инвентарный номер\n(или заводской\nномер при\nотсутствии инв.)',
    'Диапазон измерений',
    'Точность измерений',
    'Номер аттестата(свидетельства)',
    'Дата аттестации(поверки) очередной',
    '1',
    '2',
    '3',
    '4',
    '5',
    '6',
    '7',
]


list_mean_SI_IO = [
    ['Измеритель сопротивления жил кабеля', 'КИС 115', '1344', '(5х10^(-6)-170) Ом', '±(0,2-2) %', 'С-МА/26-08-2022/181236475', '25.08.2023'],
    ['Установка высоковольтная измерительная (испытательная)', 'УПУ – 21/2', '199', '(1-3-10) кВ', '±3 %', '551-73668-2022-199', '25.08.2023'],
    ['Рулетка измерительная металлическая', 'Р50У3К', 'E1152', '(0—50) м', 'кл. 3', 'С-МА/18-04-2023/239643005', '17.04.2024'],
    ['Микрометр гладкий с ценой деления 0,01 мм', 'МК 25', '3429', '(0—25) мм', '±0,002 мм', 'С-БВК/20-04-2023/241821400', '19.04.2024'],
    ['Штангенциркуль', 'ШЦЦ-I-150-0.01', 'G131103', '(0—150) мм', '±0,03 мм', 'С-БВК/20-04-2023/242645501', '19.04.2024'],
    ['Термогигрометр', 'ИВА-6Н-Д', '5022', '0—98%n\n(-20—50)°С\n(700—1100)гПа', '±2%;±3%\n±0,3°С\n±2,5гПа', 'С-МА/28-12-2022/212174005', '27.12.2023'],
    ['Тераомметр', 'ТОмМ-01', '2012-12', '(10^6-10^15) Ом', '±(5-10) %', 'С-МА/29-12-2022/212632869', '28.12.2023'],
    ['Система для измерения плотности дыма (светопроницаемости) при горении и тлении кабельного изделия', 'УИПД', '01', '-', '-', '06А-23', '20.03.2024'],
    ['Установка для испытаний электрических и оптических кабелей и проводов, проложенных пучком, на нераспространение горения', 'УИНГ-П', '1449', '-', '-', '21А-22', '28.12.2023'],
    ['Машина универсальная испытательная', 'Z010', '730534', '0,01 H—10 кН', '±1,0 % ±0,5 %', 'С-МА/30-08-2022/182700446', '29.08.2023'],
    ['Климатическая камера', 'SE-600-6-6', '1384', '(-70 - +180)°С (10 – 98)%', '±0,3 °С; ±0,5 °С ±2,5 %;  ±1,0 %', '20А-22', '12.12.2024'],
]

list_head_test_table = [
    'Наименование показателя, размерность',
    'Номера пунктов НД',
    'Согласно НД',
    'Фактическое значение показателя образца',
    'Вывод о соответствии',
    'технических требований',
    'методов испытаний',
    'Значение показателя',
    'Допуск',
]

dict_tests = {
1: 'Проверка внешнего вида маркировки, ее разборчивости и содержания', 
2: 'Испытание маркировки на прочность к воздействию влаги ',
3: 'Проверка внешнего вида кабеля',
4: 'Контроль электрического сопротивления изоляции токопроводящих жил',
5: 'Испытание напряжением постоянного тока между токопроводящими жилами пары',
6: 'Испытание напряжением постоянного тока между всеми токопроводящими жилами и экраном',
7: 'Контроль отсутствия обрывов жил, экранов, контактов между токопроводящими жилами, между токопроводящими жилами и экраном',
8: 'Проверка внешнего вида маркировки, ее разборчивости и содержания', 
9: 'Контроль электрического сопротивления токопроводящих жил',
10: 'Контроль омической асимметрии изолированных токопроводящих жил в паре',
11: 'Контроль электрической емкости пары',
12: 'Контроль емкостной асимметрии пар по отношению к экрану',
13: 'Контроль максимальной разности времени задержки сигнала',
14: 'Контроль коэффициента затухания пар',
15: 'Контроль переходного затухания суммарной мощности влияния на ближнем конце',
16: 'Контроль переходного затухания на ближнем конце',
17: 'Контроль защищенности пар на дальнем конце',
18: 'Контроль волнового сопротивления пар',
19: 'Проверка общего вида, элементов конструкции и основных размеров кабеля',
20: 'Испытания на стойкость к продольному гидростатическому давлению',
21: 'Испытания на безотказность (кратковременные) продолжительностью 500 (1000) ч',
22: 'Контроль относительного удлинения при разрыве токопроводящей жилы',
23: 'Контроль относительного удлинения при разрыве изоляции токопроводящей жилы',
24: 'Контроль относительного удлинения при разрыве оболочки',
25: 'Контроль прочности при растяжении изоляции токопроводящей жилы',
26: 'Контроль прочности при растяжении оболочки',
27: 'Контроль стойкости к многократным изгибам',
28: 'Контроль усадки линейных размеров изоляции',
29: 'Испытание на стойкость к воздействию повышенной температуры среды',
30: 'Испытание на стойкость к воздействию пониженной температуры среды',
31: 'Испытание на стойкость к воздействию изменения температуры среды',
32: 'Испытание на стойкость к воздействию повышенной влажности воздуха',
33: 'Испытание маркировки на сохранение разборчивости и прочности при эксплуатации, транспортировании и хранении',
34: 'Проверка массы',
35: 'Проверка возможности вертикальной прокладки кабеля',
36: 'Испытания на безотказность (длительные) продолжительностью 2000 ч',
37: 'Испытание на нераспространение горения при групповой прокладке',
38: 'Испытание на дымообразование',
39: 'Контроль сопротивления связи',
40: 'Контроль температурного коэффициента затухания пар',
41: 'Испытание на стойкость к воздействию солнечного излучения',
42: 'Испытания на стойкость к повышенному атмосферному давлению',
43: 'Испытания на стойкость к пониженному атмосферному давлению',
44: 'Проверка коррозийной активности продуктов дымо и газовыделения при горении',
45: 'Испытания на стойкость к гидростатическому радиальному давлению',
46: 'Испытания на стойкость к горючесмазочным материалам (ГСМ)',
47: 'Испытания на стойкость к морской воде',
48: 'Испытания на стойкость к раствору щавелевой кислоты',
49: 'Испытания кабеля на воздействие предельной повышенной температуры среды',
50: 'Испытание на воздействие атмосферных конденсированных осадков (инея и росы)',
51: 'Контроль относительного удлинения при разрыве оболочки из композиции, не содержащей галогенов после теплового старения',
52: 'Контроль прочности при растяжении оболочки из композиции, не содержащей галогенов после теплового старения',
52: 'Испытания на стойкость к продольному гидростатическому давлению в течение 24 ч',
}

list_records_test = {
    1: [
    'Электрическое сопротивление токопроводящих жил, пересчитанное на 1 км длины и температуру 20 С, Ом',
    'Электрическое сопротивление изоляции жил, пересчитанное на 1 км длины и температуру 20 С, МОм',
    'Испытание напряжением переменного тока номинальной частотой 50 Гц в течение 1 мин между жилами и между жилами и экраном, 700 В',
    'Электрическая емкость пары, пересчитанная на длину 1000 м, нФ'
    'Волновое сопротивление, Ом'
    ],
    2: [
    'Относительное удлинение при разрыве изоляции, %',
    'Прочность при разрыве изоляции, Н/мм2',
    'Относительное удлинение при разрыве оболочки, %',
    'Прочность при разрыве оболочки, Н/мм2',
    'Относительное удлинение при разрыве токопроводящей жилы, %',  
    ],
    3: [
    'Стойкость к воздействию повышенной температуры окружающей среды',
    'Стойкость к воздействию пониженной температуры окружающей среды',
    'Стойкость к воздействию повышенной влажности воздуха',
    'Стойкость к изгибу кабеля при пониженной температуре',
    'Стойкость к воздействию дизельного топлива',
    'Стойкость к воздействию минерального масла СЖР-2',
    ],
    4: [
    'Маркировка',
    'Прочность маркировки',
    'Ярлык',
    ]
}