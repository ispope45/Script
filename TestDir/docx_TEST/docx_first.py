
if __name__ == "__main__":
    text_file_path = "1.doc"
    new_text_content = ''
    target_word1 = '></SPAN>'
    new_word1 = '>800M</SPAN>'

    target_word2 = '></SPAN>'
    new_word2 = '>1</SPAN>'

    target_word3 = '>M</SPAN>'
    new_word3 = '>999M</SPAN>'

    target_word4 = '>M</SPAN>'
    new_word4 = '>888M</SPAN>'

    with open(text_file_path, 'r', encoding='utf8') as f:
        lines = f.readlines()
        for i, l in enumerate(lines):
            # print(i)
            if i == 148:
                new_string = l.strip().replace(target_word1, new_word1)
                new_text_content += new_string
                print(new_string)
            elif i == 153:
                new_string = l.strip().replace(target_word2, new_word2)
                new_text_content += new_string
                print(new_string)
            elif i == 350:
                new_string = l.strip().replace(target_word3, new_word3)
                new_text_content += new_string
                print(new_string)
            elif i == 361:
                new_string = l.strip().replace(target_word4, new_word4)
                new_text_content += new_string
                print(new_string)
            else:
                new_string = l
                if new_string:
                    new_text_content += new_string
            # print(l)
            # new_string = l.strip().replace(target_word, new_word)

    with open(text_file_path, 'wb') as f:
        res_text = bytes(new_text_content, 'utf-8')
        f.write(res_text)
