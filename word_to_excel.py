import docx
from docx.document import Document
import openpyxl
import os, re

filename = "Java考古題(選擇).docx"
doc = docx.Document(filename)
para = doc.paragraphs
print('段落數量： ', len(para),'\n')
i = 0

for _ in range(0, len(para)):
    try:
        content = para[_].text
        # print([content[0], content[4], content[5]])
        # if [content[0], content[4], content[5]] == ['(', ')', '：']: # 題目
        #     print(content)
        # if [content[0], content[1]] == ['A', '：']: # 選項A
        #     print(content)
        # if [content[0], content[1]] == ['B', '：']: # 選項B
        #     print(content)
        # if [content[0], content[1]] == ['C', '：']: # 選項C
        #     print(content)
        # if [content[0], content[1]] == ['D', '：']: # 選項D
        #     print(content)
        img = para[_]._element.xpath(".//pic:pic")
        if img:
            # i += 1
            # print(i, img)
            img = img[0]
            rel = img.xpath('.//a:blip/@r:embed')[0]
            imagepart = doc.part.related_parts[rel]
            print(imagepart)
    except IndexError:
        pass

def get_pictures(word_path, result_path): # 讀取圖片
    """
    图片提取
    :param word_path: word路径
    :return: 
    """
    try:
        doc = docx.Document(word_path)
        dict_rel = doc.part._rels
        for rel in dict_rel:
            print(rel)
            rel = dict_rel[rel]
            if "image" in rel.target_ref:
                if not os.path.exists(result_path):
                    os.makedirs(result_path)
                img_name = re.findall("/(.*)", rel.target_ref)[0]
                word_name = os.path.splitext(word_path)[0]
                if os.sep in word_name:
                    new_name = word_name.split('\\')[-1]
                else:
                    new_name = word_name.split('/')[-1]
                img_name = f'{new_name}-'+'.'+f'{img_name}'
                with open(f'{result_path}/{img_name}', "wb") as f:
                    f.write(rel.target_part.blob)
    except:
        pass

def get_word_pictures(): # 讀取圖片
    global filename
    """
    图片提取
    :param word_path: word路径
    :return: 
    """
    try:
        doc = docx.Document(filename)
        dict_rel = doc.part._rels
        print(dict_rel)
        for index, rel in enumerate(dict_rel):
            rel = dict_rel[rel]
            print(index, rel.target_ref)
            # if "image" in rel.target_ref:
            #     if not os.path.exists(result_path):
            #         os.makedirs(result_path)
            #     img_name = re.findall("/(.*)", rel.target_ref)[0]
            #     word_name = os.path.splitext(word_path)[0]
            #     if os.sep in word_name:
            #         new_name = word_name.split('\\')[-1]
            #     else:
            #         new_name = word_name.split('/')[-1]
            #     img_name = f'{new_name}-'+'.'+f'{img_name}'
            #     with open(f'{result_path}/{img_name}', "wb") as f:
            #         f.write(rel.target_part.blob)
    except:
        pass

# get_word_pictures()
# os.chdir("C:\MyProgram\考題練習")
# spam=os.listdir(os.getcwd())
# for i in spam:
#     get_pictures(str(i),os.getcwd())