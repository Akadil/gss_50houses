import docx
from docx.shared import Pt


container_4etap_3tip = [[51, 118],[61, 123], [49, 117], [55, 120], [57, 121], [65, 125], [63, 124], [53, 119], [59, 122]]
container_4etap_2tip = [[51, 99], [53, 100], [55, 101], [57, 102], [59, 103], [61, 104], [63, 105], [65, 106], [67, 107]]
container_4etap_1tip = [[52, 136], [56, 138], [50, 135], [48, 134], [46, 133], [58, 139], [54, 137], [60, 140], [62, 141]]

container_3etap_3tip = [[47, 116], [45, 115], [29, 107], [33, 109], [35, 110],
                        [27, 142], [37, 111], [31, 108], [39, 112], [43, 114], [41, 113]]
container_3etap_2tip = [[45, 96], [43, 95], [47, 97], [49, 98], [39, 93], [41, 94]]
container_3etap_1tip = [[44, 132], [34, 127], [36, 128], [40,130], [42, 131], [38, 129]]

name_bases = ['Заключение авторский 3 эт улкенбаева', 'Заключение авторский 3 эт сарыбаева',
              'Заключение авторский 3 эт сарыбаева', 'Заключение авторский 4 эт улкенбаева',
              'Заключение авторский 4 эт сарыбаева', 'Заключение авторский 4 эт сарыбаева']

number_bases = [container_3etap_3tip, container_3etap_2tip, container_3etap_1tip,
                container_4etap_3tip, container_4etap_2tip, container_4etap_1tip]

for m in range(0, 6):

    doc = docx.Document(f'{name_bases[m]}.docx')
    all_paras = doc.paragraphs

    for item in number_bases[m]:
        if len(item) == 0:
            continue
        # print(all_paras[8].text)
        # print(all_paras[11].text)
        # print("\n")
        smth = all_paras[19].text
        left_position = all_paras[19].text.find('аев') + 4
        strin = all_paras[19].text[left_position:left_position + 7]
        i = 1
        for letter in strin:
            if letter == '0' or letter == '1' or letter == '2' or letter == '3' \
                    or letter == '4' or letter == '5' or letter == '6' \
                    or letter == '7' or letter == '8' or letter == '9':
                i = i + 1
                continue
            break
        right_position = i - 1 + left_position
        # print(all_paras[8].text[left_position:right_position])
        # print(len(all_paras[8].text[left_position:right_position]))

        all_paras[19].text = all_paras[19].text[:left_position] + str(item[0]) + all_paras[19].text[right_position:]

        style = doc.styles['Normal']
        font = style.font
        font.name = 'Times New Roman'
        font.size = Pt(12)
        all_paras[19].style = doc.styles['Normal']

        if m == 0:
            doc.save(f'Нарынкол 23 ({str(item[0])} Улкенбаева) - Заключение автор.docx')
        elif m == 1 or m == 2:
            doc.save(f'Нарынкол 23 ({str(item[0])} Сарыбаева) - Заключение автор.docx')
        elif m == 3:
            doc.save(f'Нарынкол 27 ({str(item[0])} Улкенбаева) - Заключение автор.docx')
        elif m == 4 or m == 5:
            doc.save(f'Нарынкол 27 ({str(item[0])} Сарыбаева) - Заключение автор.docx')
        else:
            print("Something go wrong!")

        # doc.save(f'{name_bases[m]} ({item[0]}).docx')
        # print(all_paras[8].text)
        # print(all_paras[11].text)
        # print('\n\n')


""" 20
i = 1
for item in all_paras:
    print(f'{i}th parapgraph -> {item.text}')
    i = i + 1
"""