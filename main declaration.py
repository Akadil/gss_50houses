from docx import Document
from docx.shared import Pt


container_4etap_3tip = [[51, 118],[61, 123], [49, 117], [55, 120], [57, 121], [65, 125], [63, 124], [53, 119], [59, 122]]
container_4etap_2tip = [[51, 99], [53, 100], [55, 101], [57, 102], [59, 103], [61, 104], [63, 105], [65, 106], [67, 107]]
container_4etap_1tip = [[52, 136], [56, 138], [50, 135], [48, 134], [46, 133], [58, 139], [54, 137], [60, 140], [62, 141]]

container_3etap_3tip = [[47, 116], [45, 115], [29, 107], [33, 109], [35, 110],
                        [27, 142], [37, 111], [31, 108], [39, 112], [43, 114], [41, 113]]
container_3etap_2tip = [[45, 96], [43, 95], [47, 97], [49, 98], [39, 93], [41, 94]]
container_3etap_1tip = [[44, 132], [34, 127], [36, 128], [40,130], [42, 131], [38, 129]]

name_bases = ['Декларация 3 эт Тип 3', 'Декларация 3 эт Тип 2',
              'Декларация 3 эт Тип 1', 'Декларация 4 эт Тип 3',
              'Декларация 4 эт Тип 2', 'Декларация 4 эт Тип 1']

number_bases = [container_3etap_3tip, container_3etap_2tip, container_3etap_1tip,
                container_4etap_3tip, container_4etap_2tip, container_4etap_1tip]


for m in range(0, 6):


    doc = Document(f'{name_bases[m]}.docx')
    all_paras = doc.paragraphs

    #for s in range(0, len(all_paras)):
    #    print(f'{s}th  {all_paras[s].text}')

    for item in number_bases[m]:

        left_position = all_paras[13].text.find('аев') + 4
        strin = all_paras[13].text[left_position:left_position + 7]
        i = 1
        for letter in strin:
            if letter == '0' or letter == '1' or letter == '2' or letter == '3' \
                    or letter == '4' or letter == '5' or letter == '6' \
                    or letter == '7' or letter == '8' or letter == '9':
                i = i + 1
                continue
            break
        right_position = i - 1 + left_position

        all_paras[13].text = all_paras[13].text[:left_position] + str(item[0]) + all_paras[13].text[right_position:]

        style = doc.styles['Normal']
        font = style.font
        font.name = 'Times New Roman'
        font.size = Pt(12)
        all_paras[13].style = doc.styles['Normal']

        if m == 0:
            doc.save(f'Нарынкол 23 ({str(item[0])} Улкенбаева) - Декларация.docx')
        elif m == 1 or m == 2:
            doc.save(f'Нарынкол 23 ({str(item[0])} Сарыбаева) - Декларация.docx')
        elif m == 3:
            doc.save(f'Нарынкол 27 ({str(item[0])} Улкенбаева) - Декларация.docx')
        elif m == 4 or m == 5:
            doc.save(f'Нарынкол 27 ({str(item[0])} Сарыбаева) - Декларация.docx')
        else:
            print("Something go wrong!")
