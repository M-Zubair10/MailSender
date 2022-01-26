import csv
import ctypes
import smtplib
import ssl
import sys
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

import PySimpleGUI as sg

ctypes.windll.shcore.SetProcessDpiAwareness(2)

attachments = []


def Space(n):
    return " " * n


def write_file(email="imuhammadzubair207@gmail.com", password="appsamle10"):
    with open("credentials.csv", "w+", newline='') as file:
        writer = csv.writer(file)
        writer.writerow(["Email", "Password"])
        writer.writerow([email, password])


def read_file(header=False):
    rows = []
    with open("credentials.csv", 'r') as file:
        csvreader = csv.reader(file)
        h = next(csvreader)
        for row in csvreader:
            rows.append(row)
    if header:
        return h, rows
    return rows


def write_receiver_addresses(value):
    with open("receivers.txt", "a+") as file:
        return file.writelines(value + "\n")


def read_receiver_addresses():
    with open("receivers.txt", "r+") as file:
        return file.readlines()


def get_theme():
    """
    Get the theme to use for the program
    Value is in this program's user settings. If none set, then use PySimpleGUI's global default theme
    :return: The theme
    :rtype: str
    """
    try:
        global_theme = sg.theme_global()
    except:
        global_theme = sg.theme()
    user_theme = sg.user_settings_get_entry('-theme-', '')
    if user_theme == '':
        user_theme = global_theme
    return user_theme


def make_settings():
    theme = get_theme()
    if not theme:
        theme = sg.OFFICIAL_PYSIMPLEGUI_THEME
    sg.theme(theme)
    event, values = sg.Window("Settings", [
        [sg.T('Theme', font='_ 16')],
        [sg.T('Leave blank to use global default'), sg.T(sg.theme_global())],
        [sg.Combo([''] + sg.theme_list(), sg.user_settings_get_entry('-theme-', ''), readonly=True, k='-THEME-')],
        [sg.B("Save", size=(8, 1)), sg.B("Cancel", size=(8, 1))]
    ]).read(close=True)
    if event == "Save" and values["-THEME-"] != sg.theme_global():
        sg.user_settings_set_entry('-theme-', values['-THEME-'])
        return True


def make_window():
    global attachments
    theme = get_theme()
    if not theme:
        theme = sg.OFFICIAL_PYSIMPLEGUI_THEME
    sg.theme(theme)
    sg.theme("DefaultNoMoreNagging")
    layout = [
        [sg.T("Credentials", font=("Ariel", 20)), sg.B(" ", visible=False, disabled=True)],
        [sg.T("Sender Address: {}".format(Space(2))),
         sg.B(read_file()[0][0], key="-SENDER-", size=(35, 1), button_color=sg.theme_background_color())],
        [sg.T("Password: {}".format(Space(11))),
         sg.B("●" * len(read_file()[0][1]), key="-PASS-", size=(35, 1), button_color=sg.theme_background_color())],
        [sg.T("Receiver Address: "),
         sg.Listbox(" " if read_receiver_addresses() == [] else read_receiver_addresses(), size=(40, 1),
                    font=("Ariel", 11),
                    key="-RECEIVE-", no_scrollbar=True, )],
        [sg.T(Space(44)), sg.B("Add", size=(6, 1)), sg.B("Remove"), sg.B("Remove All")],
        [sg.HorizontalSeparator()],
        [sg.T("Email", font=("Ariel", 20))],
        [sg.Frame('', [
            [sg.T("Subject: " + Space(2)), sg.InputText(key="-SUBJECT-")],
            [sg.T("Message ")],
            [sg.Multiline(size=(50, 10), key="-OUTPUT-")],
            [sg.FileBrowse("Attachments"),
             sg.B('B', font=("Ariel", 10, 'bold'), button_color=sg.theme_background_color(), size=(3, 1)),
             sg.B('I', font=("Ariel", 10, 'italic'), button_color=sg.theme_background_color(), size=(3, 1)),
             sg.B('U', font=("Ariel", 10, 'underline'), button_color=sg.theme_background_color(), size=(3, 1))]
        ])],
        [sg.B("Send", size=(10, 1)),
         sg.B('MultiSend', button_color=sg.theme_background_color(), size=(15, 1)),
         sg.T(Space(5)), sg.B("Settings", size=(10, 1)), sg.B("Exit", size=(8, 1))]
    ]

    window = sg.Window("Email", layout, size=(600, 720))
    prev_value = None
    while True:
        event, value = window.read(timeout=100)
        if event in (sg.WINDOW_CLOSED, "Exit"):
            window.close()
            sys.exit()
        elif event == "Send":
            send_mail(value)
        elif event == "MultiSend":
            sub_event, sub_value = sg.Window("MultiSend", [
                [sg.T("How many times you want to send this message")],
                [sg.Input(key="-IN-")],
                [sg.Ok(size=(8, 1)), sg.B("Cancel", size=(8, 1))]
            ]).read(close=True)
            if sub_event == "Ok":
                send_mail(value, int(sub_value["-IN-"]))
        elif event == "-SENDER-":
            sub_event, sub_value = sg.Window("New Sender", [
                [sg.T("Add new sending address")],
                [sg.Input(key="-INPUT-")],
                [sg.B("Ok", size=(6, 1)), sg.B("Cancel")]
            ], keep_on_top=True).read(close=True)
            if sub_event == "Ok" and not sub_value["-INPUT-"] == '':
                write_file(sub_value["-INPUT-"], read_file()[0][1])
        elif event == "-PASS-":
            sub_event, sub_value = sg.Window("New Password", [
                [sg.T("Add new password")],
                [sg.Input(key="-INPUT-")],
                [sg.B("Ok", size=(6, 1)), sg.B("Cancel")]
            ], keep_on_top=True).read(close=True)
            if sub_event == "Ok" and not sub_value["-INPUT-"] == '':
                write_file(email=read_file()[0][0], password=sub_value["-INPUT-"])
        elif event == "Add":
            sub_event, sub_value = sg.Window("New Receiver", [
                [sg.T("Add new receiving address")],
                [sg.Input(key="-INPUT-")],
                [sg.B("Ok", size=(6, 1)), sg.B("Cancel")]
            ], keep_on_top=True).read(close=True)
            if sub_event == "Ok":
                write_receiver_addresses(sub_value["-INPUT-"])
        elif event == "Remove":
            temp = read_receiver_addresses()
            try:
                temp.pop()
            except IndexError:
                continue
            with open("receivers.txt", "w") as file:
                file.writelines(temp)
        elif event == "Remove All":
            with open("receivers.txt", "w") as file:
                file.writelines("")
        elif event == "Settings":
            setting = make_settings()
            if setting:
                window.close()
                make_window()
        elif event == "B":
            window["-OUTPUT-"].update(value["-OUTPUT-"] + "** **")
        elif event == "I":
            window["-OUTPUT-"].update(value["-OUTPUT-"] + "` `")
        elif event == "U":
            window["-OUTPUT-"].update(value["-OUTPUT-"] + "! !")

        if not prev_value == value["Attachments"] and not value["Attachments"] == "":
            attachments.append(value["Attachments"])
        prev_value = value["Attachments"]

        window["-SENDER-"].update(read_file()[0][0])
        window["-PASS-"].update("●" * len(read_file()[0][1]))
        window["-RECEIVE-"].update(read_receiver_addresses())


def send_mail(value, n=1):
    orig_message = ""
    sender_email = read_file()[0][0]
    receiver_email = read_receiver_addresses()
    password = read_file()[0][1]
    message = MIMEMultipart("alternative")
    message["Subject"] = value["-SUBJECT-"]
    message["From"] = sender_email
    message["To"] = "\n".join(receiver_email)

    for i in value["-OUTPUT-"].split(" "):
        if i == "**" or i == "` `" or i == "! !":
            continue
        elif "`" in i:
            i = list(i)
            i[0] = "<i>"
            i[-1] = "</i>"
            i = "".join(i)
        elif "!" == i[0] and "!" == i[-1]:
            i = list(i)
            i[0] = "<u>"
            i[-1] = "</u>"
            i = "".join(i)
        elif "**" in i:
            i = list(i)
            i[0] = "<b>"
            i.pop(1)
            i[-1] = "</b>"
            i.pop(-2)
            i = "".join(i)
        orig_message += i

    if attachments:
        for file in attachments:
            # Open files in binary mode
            with open(file, "rb") as attachment:
                # Add file as application/octet-stream
                # Email client can usually download this automatically as attachment
                part = MIMEBase("application", "octet-stream")
                part.set_payload(attachment.read())

            # Encode file in ASCII characters to send by email
            encoders.encode_base64(part)

            # Add header as key/value pair to attachment part
            part.add_header(
                "Content-Disposition",
                f"attachment; filename= {file}",
            )

            # Add attachment to message and convert message to string
            message.attach(part)

    # Create the plain-text and HTML version of your message
    text = """
    Hello
    """
    html = """\
    <html>
      <body>
        <p>{}
        </p>
      </body>
    </html>
    """.format(str(orig_message))

    # Turn these into plain/html MIMEText objects
    part1 = MIMEText(text, "plain")
    part2 = MIMEText(html, "html")

    # Add HTML/plain-text parts to MIMEMultipart message
    # The email client will try to render the last part first
    message.attach(part1)
    message.attach(part2)

    # Create secure connection with server and send email
    context = ssl.create_default_context()
    with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as server:
        server.login(sender_email, password)
        for i in range(n):
            server.sendmail(sender_email, receiver_email, message.as_string())


if __name__ == '__main__':
    make_window()
