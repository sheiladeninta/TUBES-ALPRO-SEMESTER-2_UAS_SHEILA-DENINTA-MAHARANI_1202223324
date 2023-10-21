from docxtpl import DocxTemplate

doc = DocxTemplate("invoice_template.docx")

invoice_list = [[2, "pen", 0.5, 1],
                [1, "paper pack", 5, 5],
                [2, "notebook", 2, 4]]


doc.render({"name":"john", 
            "phone":"555-55555",
            "tgl_order": "27/01/2023",
            "tgl_jatuhtempo": "27/02/2023",
            "no_invoice": "140",
            "invoice_list": invoice_list,
            "subtotal":10,
            "salestax":"10%",
            "total":9})
doc.save("new_invoice.docx")
