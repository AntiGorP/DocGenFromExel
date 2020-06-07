from docxtpl import DocxTemplate



def docsCreator(tempfile, resfile, data):

    # Генерируем документы по шаблону
    
    docf = DocxTemplate(tempfile)
    docf.render(data)
    docf.save(resfile)

