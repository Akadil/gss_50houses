import docx

container_4etap_3tip = [[51, 118],[61, 123], [49, 117], [55, 120], [57, 121], [65, 125], [63, 124], [53, 119], [59, 122]]
container_4etap_2tip = [[51, 99], [53, 100], [55, 101], [57, 102], [59, 103], [61, 104], [63, 105], [65, 106], [67, 107]]
container_4etap_1tip = [[52, 136], [56, 138], [50, 135], [48, 134], [46, 133], [58, 139], [54, 137], [60, 140], [62, 141]]

container_3etap_3tip = [[47, 116], [45, 115], [29, 107], [33, 109], [35, 110],
                        [27, 142], [37, 111], [31, 108], [39, 112], [43, 114], [41, 113]]
container_3etap_2tip = [[45, 96], [43, 95], [47, 97], [49, 98], [39, 93], [41, 94]]
container_3etap_1tip = [[44, 132], [34, 127], [36, 128], [40,130], [42, 131], [38, 129]]


name_bases = ['-Нарынкол 23 (27 Улкенбаева) - Тип 3.docx', '-Нарынкол 23 (43 Сарыбаева) - Тип 2.docx',
              '-Нарынкол 23 (44 Сарыбаева) - Тип 1.docx', '-Нарынкол 27 (49 Улкенбаева) - 4 этап Тип 3.docx',
              '-Нарынкол 27 (49 Сарыбаева) - 4 этап Тип 2.docx', '-Нарынкол 27 (46 Сарыбаева) - 4 этап Тип 1.docx']
number_bases = [container_3etap_3tip, container_3etap_2tip, container_3etap_1tip,
                container_4etap_3tip, container_4etap_2tip, container_4etap_1tip]

for m in range(0, 6):

    doc = docx.Document(name_bases[m])
    all_paras = doc.paragraphs

    for item in number_bases[m]:

        left_position_num = all_paras[12].text.find('№') + 2
        strin = all_paras[12].text[left_position_num:left_position_num + 7]
        i = 1
        for letter in strin:
            if letter == ' ':
                break
            i = i + 1
        right_position_num = i - 1 + left_position_num

        left_position = all_paras[9].text.find('аев') + 4
        strin = all_paras[9].text[left_position:left_position + 5]
        i = 1
        for letter in strin:
            if letter == ',':
                break
            i = i + 1
        right_position = i - 1 + left_position

        all_paras[9].text = all_paras[9].text[:left_position] + str(item[0]) + all_paras[9].text[right_position:]
        all_paras[12].text = all_paras[12].text[:left_position_num] + str(item[1]) + all_paras[12].text[right_position_num:]

        if m == 0:
            doc.save(f'Нарынкол 23 ({str(item[0])} Улкенбаева) - Акт приемки.docx')
        elif m == 1 or m == 2:
            doc.save(f'Нарынкол 23 ({str(item[0])} Сарыбаева) - Акт приемки.docx')
        elif m == 3:
            doc.save(f'Нарынкол 27 ({str(item[0])} Улкенбаева) - Акт приемки.docx')
        elif m == 4 or m == 5:
            doc.save(f'Нарынкол 27 ({str(item[0])} Сарыбаева) - Акт приемки.docx')
        else:
            print("Something go wrong!")
