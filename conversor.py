try:
    import os
    import fitz  # PyMuPDF
    from colorama import Fore, init
    from pdf2docx import Converter
    from docx import Document
    from reportlab.lib.pagesizes import letter
    from reportlab.pdfgen import canvas
except ImportError as e:
    print(f"Error al importar una librería: {e}")
init(autoreset=True)
def mostrar_banner():
    banner = """
    _ (`-.  _ .-') _                                _  .-')         _ .-') _                         ) (`-.      
  ( (OO  )( (  OO) )                              ( \( -O )       ( (  OO) )                         ( OO ).    
 _.`     \ \     .'_    ,------.       .-'),-----. ,------.        \     .'_  .-'),-----.    .-----.(_/.  \_)-. 
(__...--'' ,`'--..._)('-| _.---'      ( OO'  .-.  '|   /`. '       ,`'--..._)( OO'  .-.  '  '  .--./ \  `.'  /  
 |  /  | | |  |  \  '(OO|(_\          /   |  | |  ||  /  | |       |  |  \  '/   |  | |  |  |  |('-.  \     /\  
 |  |_.' | |  |   ' |/  |  '--.       \_) |  |\|  ||  |_.' |       |  |   ' |\_) |  |\|  | /_) |OO  )  \   \ |  
 |  .___.' |  |   / :\_)|  .--'         \ |  | |  ||  .  '.'       |  |   / :  \ |  | |  | ||  |`-'|  .'    \_) 
 |  |      |  '--'  /  \|  |_)           `'  '-'  '|  |\  \        |  '--'  /   `'  '-'  '(_'  '--'\ /  .'.  \  
 `--'      `-------'    `--'               `-----' `--' '--'       `-------'      `-----'    `-----''--'   '--' 
    """
    print(Fore.LIGHTRED_EX + banner)
    print(Fore.LIGHTRED_EX + "PDF to DOCX converter or Docx to PDF converter")
    print(Fore.LIGHTRED_EX + "Author: @dev-Thomas")
    print(Fore.LIGHTRED_EX + "Version: 1.0")

def pdf_a_word(ruta_pdf):
    ruta_word = ruta_pdf.replace('.pdf', '.docx')
    cv = Converter(ruta_pdf)
    cv.convert(ruta_word, start=0, end=None)
    cv.close()
    print(f"Archivo convertido a Word: {ruta_word}")

def word_a_pdf(ruta_word):
    ruta_pdf = ruta_word.replace('.docx', '.pdf')
    doc = Document(ruta_word)
    c = canvas.Canvas(ruta_pdf, pagesize=letter)
    width, height = letter

    for para in doc.paragraphs:
        c.drawString(72, height - 72, para.text)
        c.showPage()

    c.save()
    print(f"Archivo convertido a PDF: {ruta_pdf}")

def main():
    mostrar_banner()
    while True:
        print("1. Convertir PDF a Word")
        print("2. Convertir Word a PDF")
        print("3. Salir")
        opcion = input("Elige una opción: ")

        if opcion == '1':
            ruta_pdf = input("Ingresa la ruta del archivo PDF: ")
            try:
                if os.path.exists(ruta_pdf):
                    pdf_a_word(ruta_pdf)
                else:
                    print("El archivo no existe. Intenta nuevamente.")
            except:
                print(f"ERROR al ingresar a <<{ruta_pdf} >> Intentelo nuevamente.")
        elif opcion == '2':
            ruta_word = input("Ingresa la ruta del archivo Word: ")
            if os.path.exists(ruta_word):
                word_a_pdf(ruta_word)
            else:
                print("El archivo no existe. Intenta nuevamente.")
        elif opcion == '3':
            print("Saliendo...")
            break
        else:
            print("Opción no válida. Intenta nuevamente.")

if __name__ == "__main__":
    main()