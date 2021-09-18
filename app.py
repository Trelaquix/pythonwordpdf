from flask import Flask, request, render_template
import os
import fitz
import smtplib
import pathlib
from email import encoders
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
import sys
# import pypyodbc as odbc
import textwrap
import pyodbc
import datetime
import pypandoc
from mailmerge import MailMerge

app = Flask(__name__)

# os.chdir("C:\\Users\\bryanlon\\Desktop\\python_word_to_pdf_editor-main")
current_dir = os.path.dirname(os.path.realpath(__file__))
# print(current_dir)
# working_dir = os.path.join(current_dir , "..")

try:
    #Specify the driver
    driver = '{ODBC Driver 17 for SQL Server}'

    #Specify the Server name and Database Name
    server_name = 'pythonwordpdf'
    database_name = 'wordpdf'

    server = '{server_name}.database.windows.net,1433'.format(server_name=server_name)
    print("1.1")
    # Username & password
    username = "bryan_admin"
    password = "j2W3j4@qdY7YZ8V"

    # Create connection string
    connection_string = textwrap.dedent('''
        Driver={driver};
        Server={server};
        Database={database};
        Uid={username};
        Pwd={password};
        Encrypt=yes;
        TrustServerCertificate=no;
        Connection Timeout=30;
    '''.format(
        driver=driver,
        server=server,
        database=database_name,
        username=username,
        password=password
    ))
    print("1.2")
    print(connection_string)
    # Create a new PYODBC Connection object
    cnxn: pyodbc.Connection = pyodbc.connect(connection_string)
    print("1.3")

    # Create a new Cursor Object from the connection
    crsr: pyodbc.Cursor = cnxn.cursor()
    print("1.4")

    # Close the connection
    cnxn.close()
    print("Connected to database"+database_name+"successfully")
except:
    print("Unable to connect to database")




# getting the merge fields from the document
template = "draft.docx"
document = MailMerge(template)
field_name = (document.get_merge_fields())

details = []

date_now = datetime.datetime.now()
date_formated = date_now.strftime("%d/%m/%Y")
date_for_sql = date_now.strftime('%Y-%m-%d')
print("1")
# this method was initially used for comtypes.client but its not working now 
# def convert_to_pdf():
#     format_code = 17

#     file_input = pathlib.Path().resolve() / "edited_draft.docx"
#     file_output = pathlib.Path().resolve() / "edited_draft.pdf"

#     # file_input = os.path.join(current_dir, "edited_draft" + "." + "docx")
#     # file_output = os.path.join(current_dir, "edited_draft" + "." + "pdf")

#     #  = (edited_word_directory)
#     #  = (edited_pdf_directory)

#     word_app = comtypes.client.CreateObject('Word.Application')
#     word_file = word_app.Documents.Open(file_input)
#     word_file.SaveAs(file_output,FileFormat=format_code)
#     word_file.Close()
#     word_app.Quit()

# Encrypting the pdf with password
def encrypt_pdf(pdf, password, outfile):
    perm = int(
                fitz.PDF_PERM_ACCESSIBILITY
                | fitz.PDF_PERM_PRINT
                | fitz.PDF_PERM_COPY
                | fitz.PDF_PERM_ANNOTATE 
            )

    encrypt_meth = fitz.PDF_ENCRYPT_AES_256
    pdf.save(outfile, encryption=encrypt_meth, user_pw=password, permissions=perm)


# Sending out the encrypted email
def send_email(rec_email, file_name):
    sender_email = "devbryantest@gmail.com"
    rec_email = rec_email
    password = "devBryantest123"
    subject = "This message was sent with python"
    content = "Download your attachment now"

    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = rec_email
    msg['Subject'] = subject
    body = MIMEText(content, 'plain')
    msg.attach(body)

    with open(file_name, "rb") as opened:
        openedfile = opened.read()
    attachedfile = MIMEApplication(openedfile, _subtype = "pdf", _encoder = encoders.encode_base64)
    attachedfile.add_header('content-disposition', 'attachment', filename = "protected.pdf")
    msg.attach(attachedfile)

    server = smtplib.SMTP('smtp.gmail.com:587')
    server.ehlo()
    server.starttls()
    server.login(sender_email, password)
    server.send_message(msg, from_addr=sender_email, to_addrs=[rec_email])

@app.route('/')
def home():
    return render_template('index.html')

@app.route('/process-form', methods=["GET", "POST"])
def login():
    if request.method == "POST":
        # Getting the values
        e_reference = request.form["e_reference"]
        # generated_date = request.form["generated_date"]
        # application_date = request.form["application_date"]
        full_name = request.form["full_name"]
        nric = request.form["nric"]
        phone = request.form["phone"]
        email = request.form["email"]
        address = request.form["address"]
        allow_post = request.form.get("allow_post")
        if allow_post == "on":
            allow_post = "X"
        else:
            allow_post = "  "
        allow_email = request.form.get("allow_email")
        if allow_email == "on":
            allow_email = "X"
        else:
            allow_email = "  "
        allow_call = request.form.get("allow_call")
        if allow_call == "on":
            allow_call = "X"
        else:
            allow_call = "  "
        allow_text = request.form.get("allow_text")
        if allow_text == "on":
            allow_text = "X"
        else:
            allow_text = "  "




        # if you want to get all checkboxes as an array, you can use this method but all the names of the checkboxes have to be the same, in this case I renamed all the checkboxes to just checkbox
        # checkbox = request.form.getlist("checkbox")

        # merging the values from the form
        
        nric_first_letter = nric[0]
        nric_last_four = nric[5:9]
        formatted_nric = nric_first_letter+"XXXX"+nric_last_four

        document.merge(
            e_reference=e_reference, 
            generated_date=date_formated, 
            application_date=date_formated, 
            full_name=full_name,
            nric=formatted_nric,
            phone=phone,
            email=email,
            address=address,
            allow_post=allow_post,
            allow_email=allow_email,
            allow_call=allow_call,
            allow_text=allow_text
            )
        
        # writing the values into the word doc and renaming it
        document.write('edited_draft.docx')

        # Appending the values to details array to display on result.html
        details.append(f"{e_reference} {date_formated} {date_formated} {full_name} {nric} {phone} {email} {address} {allow_post} {allow_email} {allow_call} {allow_text} ")
        print("2")
        # use this append structure if you're using array checkbox
        # details.append(f"{e_reference} {generated_date} {application_date} {full_name} {nric} {phone} {email} {address} {checkbox} ")

        # this method was initially used for comtypes.client but its not working now 
        # convert_to_pdf()

        # file_input = pathlib.Path().resolve() / "edited_draft.docx"
        # file_output = pathlib.Path().resolve() / "edited_draft.pdf"
        
        # try:
        #     format_code = 17
        #     file_input = os.path.join(current_dir, "edited_draft" + "." + "docx")
        #     file_output = os.path.join(current_dir, "edited_draft" + "." + "pdf")
        #     word_app = comtypes.client.CreateObject('Word.Application')
        #     word_file = word_app.Documents.Open(file_input)
        #     word_file.SaveAs(file_output,FileFormat=format_code)
        #     word_file.Close()
        #     word_app.Quit()
        # except:
        #     print("Conversion Not successful")
        #  = (edited_word_directory)
        #  = (edited_pdf_directory)

  

        # edited_word_directory = current_dir+"\edited_draft.docx"
        # edited_pdf_directory = current_dir+"\edited_draft.pdf"

        # edited_word_directory = os.path.join(current_dir+"edited_draft.docx")
        # edited_pdf_directory = os.path.join(current_dir+"edited_draft.pdf")

        # edited_word_directory = os.path.join(current_dir, "edited_draft" + "." + "docx")
        # edited_pdf_directory = os.path.join(current_dir, "edited_draft" + "." + "pdf")

        # edited_word_directory = pathlib.Path().resolve() / "edited_draft.docx"
        # edited_pdf_directory = pathlib.Path().resolve() / "edited_draft.pdf"

        # docx2pdf function to convert word to pdf
        # convert(edited_word_directory,edited_pdf_directory)
        print("3")
        pypandoc.convert_file('edited_draft.docx', 'latex', outputfile="edited_draft.pdf")
        file = 'edited_draft.pdf'
        pdf = fitz.open(file)
        encrypt_pdf(pdf, '12345', 'protected_draft.pdf')

        send_email(email, "protected_draft.pdf")
        print("4")

        # redirect to result.html once everything is done
        return render_template("result.html", details=details)

    else:
        return render_template("index.html")

if __name__ == "__main__":
    app.run(debug=True)


