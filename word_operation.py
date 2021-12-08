from docx import Document
import re
import os

def normalizeFile(path):
    fileList = os.listdir(path)
    print(f"[INFO] files name before normalization are: {fileList}")
    for filename in fileList:
        tmp_name = filename
        #del all blank space
        filename=filename.replace(' ','_')
        os.rename(path+'/'+tmp_name,path+'/'+filename)
    fileList = os.listdir(path)
    print(f'[INFO] files name before normalization are: {fileList}')


def replaceText(document,content,tex):
    all_tables = document.tables
    for tables in all_tables:
        for row in tables.rows:
            #print(row)
            for icell,cell in enumerate(row.cells):
                #print(cell.text)
                if cell.text.find(f"{content}")!=-1:
                    #print(f"loc is {icell} objet is 13641788003")
                    cell.text = f'{tex}'

    all_paragraphs = document.paragraphs
    for paragraphs in all_paragraphs:
        if(paragraphs.text.find(f"{content}")!=-1):
            newtext = re.sub(f"{content}",f'{tex}',paragraphs)
            paragraphs.text = newtext

def do_replace(dir,content,text):
    normalizeFile(dir)
    fileList = os.listdir(dir)
    for file in fileList:
        document = Document(dir+'/'+file)
        replaceText(document,content,text)
        document.save(dir+'/'+file.split('.docx')[0]+'_mod.docx')
        print(f"===================[INFO] replace {content} by {text} with {file} successfully!!!===========================")
        os.remove(dir+'/'+file)


if __name__ == "__main__":
    print("====================Welcome To Word Operation Script.=============================================================")
    print("====================Right now,We can replace test for you with all .doxc files which are put in a directory only==")
    print("====================More functions are coming soon !!!============================================================")

    import argparse
    parser = argparse.ArgumentParser('word operation')
    parser.add_argument('--dir', '-d',default='word', help='Path to the base directory of docx store.')
    parser.add_argument('--content', '-c',default='', help='contents that you want to be replaced.')
    parser.add_argument('--text', '-t',default='t', help='New contents that you want')
    args = parser.parse_args()

    do_replace(args.dir,args.content,args.text)
