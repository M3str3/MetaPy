##################################################################
#TODO                   Header      
##################################################################

author = "Nacho Mestre"
version = 1.0
header = f"""
 __       __              __                _______            
/  \     /  |            /  |              /       \           
$$  \   /$$ |  ______   _$$ |_     ______  $$$$$$$  | __    __ 
$$$  \ /$$$ | /      \ / $$   |   /      \ $$ |__$$ |/  |  /  |
$$$$  /$$$$ |/$$$$$$  |$$$$$$/    $$$$$$  |$$    $$/ $$ |  $$ |
$$ $$ $$/$$ |$$    $$ |  $$ | __  /    $$ |$$$$$$$/  $$ |  $$ |
$$ |$$$/ $$ |$$$$$$$$/   $$ |/  |/$$$$$$$ |$$ |      $$ \__$$ |
$$ | $/  $$ |$$       |  $$  $$/ $$    $$ |$$ |      $$    $$ |
$$/      $$/  $$$$$$$/    $$$$/   $$$$$$$/ $$/        $$$$$$$ |
                                                     /  \__$$ |
                                                     $$    $$/ 
                                                      $$$$$$/  
Author: {author}
Version: {version}
"""

#######################################################################################################################################################
#TODO                                                           Extensions
#######################################################################################################################################################
#TODO PDF
def pdf(fileName):
    from PyPDF2 import PdfFileReader
    pdfFile = PdfFileReader(open(fileName, 'rb'))
    meta = pdfFile.getDocumentInfo()
    for metaItem in meta:
        print(f'    {metaItem[1:]}: {meta[metaItem]}')
    return meta

#TODO DOCX

# READ
def docx_read_meta(fileName):
    import docx
    doc = docx.Document(fileName)
    metadata = {}
    properties = doc.core_properties
    print(f"""
    Version: {properties.version}
    Title: {properties.title}
    Autor: {properties.author}
    Created: {properties.created}
    Category: {properties.category}
    Identifier: {properties.identifier}
    Comments: {properties.comments}
    Content_status: {properties.content_status}
    Keyword: {properties.keywords}
    last_modified_by: {properties.last_modified_by}
    language: {properties.language}
    modified: {properties.modified}
    subject: {properties.subject}
    """)
    metadata["category"] = properties.category
    metadata["comments"] = properties.comments
    metadata["content_status"] = properties.content_status
    metadata["created"] = properties.created
    metadata["identifier"] = properties.identifier
    metadata["keywords"] = properties.keywords
    metadata["last_modified_by"] = properties.last_modified_by
    metadata["language"] = properties.language
    metadata["modified"] = properties.modified
    metadata["subject"] = properties.subject
    metadata["title"] = properties.title
    metadata["version"] = properties.version
    return metadata

#Write
def docx_edit_meta(fileName):
    import docx
    doc = docx.Document(fileName)
    properties = doc.core_properties
    print("[*] Left empty to stay with same value")

    c = input(f"Author: ({properties.author})")
    if c != '':
        doc.core_properties.author = c

    c = input(f"Category: ({properties.category})")
    if c != '':
        doc.core_properties.category = c

    c = input(f"Comments: ({properties.comments})")
    if c != '':
        doc.core_properties.comments = c

    c = input(f"Content_status: ({properties.content_status})")
    if c != '':
        doc.core_properties.content_status = c
    
    c = input(f"Created: ({properties.created})")
    if c != '':
        doc.core_properties.created = c

    c = input(f"Identifier: ({properties.identifier})")
    if c != '':
        doc.core_properties.identifier = c

    c = input(f"Keywords: ({properties.keywords})")
    if c != '':
        doc.core_properties.keywords = c

    c = input(f"Last modified by: ({properties.last_modified_by})")
    if c != '':
        doc.core_properties.last_modified_by = c

    c = input(f"Language: ({properties.language})")
    if c != '':
        doc.core_properties.language = c

    c = input(f"Modified: ({properties.modified})")
    if c != '':
        doc.core_properties.modified = c

    c = input(f"Subject: ({properties.subject})")
    if c != '':
        doc.core_properties.subject = c

    c = input(f"Title: ({properties.title})")
    if c != '':
        doc.core_properties.title = c

    c = input(f"Version: ({properties.version})")
    if c != '':
        doc.core_properties.version = c
    # Deleting the old One and replacing for new One.
    import os
    os.remove(fileName)
    doc.save(fileName)

#CLEAR
def docx_clear_meta(fileName):
    import docx
    #XXX
    from datetime import datetime
    d = datetime.now()
    doc = docx.Document(fileName)
    c=''
    doc.core_properties.author = c
    doc.core_properties.category = c
    doc.core_properties.comments = c
    doc.core_properties.content_status = c
    doc.core_properties.created = d
    doc.core_properties.identifier = c
    doc.core_properties.keywords = c
    doc.core_properties.last_modified_by = c
    doc.core_properties.language = c
    doc.core_properties.modified = d
    doc.core_properties.subject = c
    doc.core_properties.title = c
    doc.core_properties.version = c
    from os import remove as rm
    rm(fileName)
    doc.save(fileName)

# Just To print empyt lines an clean the console
def br(n_lines):
    for l in range(1,n_lines):
        print()

def get_Extension(filePath):
    import os 
    filename, file_ext =  os.path.splitext(filePath)
    return file_ext

if __name__ == '__main__':
    print(header)
    loop = True;
    while loop:
        print(header)
        print('------------------')
        print("1. View Metadata")
        print("2. Edit Metadata")
        print("3. Clear Metadata")
        print('------------------')
        br(2)
        option = input("SELECT ANY OPTION -> ")

        #TODO PREVIEW
        if (option=='1'):
            item = input("The file to preview metainfo -> ")
            extension = get_Extension(item)
            if (extension == '.pdf'):
                pdf(item)
            elif (extension == '.docx'):
                docx_read_meta(item)
            else:
                print("Invalid format ....")
            input('Press some key to continue.....')
            br(30)

        #TODO EDIT
        elif (option=='2'):
            item = input("The file to edit metainfo -> ")
            extension = get_Extension(item)
            if (extension == '.docx'):
                metadata = docx_edit_meta(item)
                docx_read_meta(item)
            
        #TODO CLEAR
        elif (option=='3'):
            item = input("[WARNING] File path to remove metainfo -> ")
            extension = get_Extension(item)
            if (extension == '.docx'):
                docx_clear_meta(item)
        elif (option == '00'):
            exit(0)
        else:
            print("xD")
        