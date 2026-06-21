# Core of the TokenLook tool
import os
import sys
import base64
import json
import time
from datetime import datetime
import tokenlook
from pathlib import Path
from tabulate import tabulate

# For colored text in terminal output
bgreen = '\033[1;32m'
bred = '\033[1;31m'
noclr = '\033[0m'
bblue = '\033[1;34m'
cyan = '\033[0;36m'
orng = '\033[0;33m'

# dedicated logs TXT file path
log_file = "log.txt"
config_file = "tokenlook_config.json"

# current tool version
VERSION_MAIN = "1.0"

send_email_banner = rf"""{bgreen}
   ____            __    ____           _ __  
  / __/__ ___  ___/ /___/ __/_ _  ___ _(_) /  
 _\ \/ -_) _ \/ _  /___/ _//  ' \/ _ `/ / /   
/___/\__/_//_/\_,_/   /___/_/_/_/\_,_/_/_/                                    
{noclr}
"""

# responsible disclosure - only this menu function below is AI generated/tweaked
def menu(banner: str, options: dict):
    """
    options format:
    {
        "option name1": function1,
        "option name2": function2
    }
    """

    while True:
        os.system("cls" if os.name == "nt" else "clear")

        print(banner)

        option_map = {
            str(i): (name, func)
            for i, (name, func) in enumerate(options.items(), start=1)
        }

        rows = [[num, name] for num, (name, _) in option_map.items()]
        rows.append(["0", "Exit"])

        print(tabulate(rows, headers=["#", "Action"], tablefmt="grid"))

        choice = input("\nSelect option: ").strip()

        if choice == "0":
            break

        if choice not in option_map:
            input("\nInvalid option. Press Enter to continue...")
            continue

        try:
            option_map[choice][1]()
        except Exception as e:
            print(f"\nError: {e}")

        input("\nPress Enter to continue...")



# Log processing and JSON processing class
class ProcessLogsAndJson:
    def __init__(self, tokenlook_log_file, json_file):
        self.tokenlook_log_file = tokenlook_log_file
        self.json_file = json_file

    # Method to write total tool logs to a TXT log file
    def write_tool_log(self, log_to_write):
        with open(self.tokenlook_log_file, 'a') as f:
            f.write(log_to_write + '\n')

    def write_json_data(self, data):
        """
        :param data: needs to be a dictionary with appropriate keys and values
        :return: nothing
        """
        with open(self.json_file, "w") as jf:
            json.dump(data, jf)

    def return_json_data(self):
        with open(self.json_file, "r") as jf:
            data = json.load(jf)
            return data

    def file_exists(self):
        """
        Check whether a file exists.
        Returns:
            bool: True if the file exists, False otherwise.
        """
        return os.path.isfile(self.json_file)

# function to get current timestamp
def get_current_timestamp():
    return datetime.now().isoformat()

# function to save JSON response
def save_json_resp(file_path, json_data):
    with open(file_path, "w", encoding="utf-8") as f:
        json.dump(json_data, f, indent=4, ensure_ascii=False)

# function to save mail body HTML in HTML file.
def save_mail_html(file_path, html_content):
    path = Path(file_path)
    # Create directories only if they don't already exist
    path.parent.mkdir(parents=True, exist_ok=True)
    # Overwrite the HTML file if it exists
    path.write_text(html_content, encoding="utf-8")
    return True

def get_html_file_content(file_path):
    if not os.path.isfile(file_path):
        raise FileNotFoundError(f"[{bred}!{noclr}]HTML file not found: {file_path}")
    with open(file_path, "r", encoding="utf-8") as file:
        html_content = file.read()
    return html_content

def save_file_bytes(file_name, directory, bytes_content):
    try:
        # Ensure directory exists (does NOT overwrite it)
        os.makedirs(directory, exist_ok=True)
        # Decode base64 string to raw bytes
        file_bytes = base64.b64decode(bytes_content)
        # Full file path
        file_path = os.path.join(directory, file_name)
        # Write file (overwrites if already exists)
        with open(file_path, "wb") as f:
            f.write(file_bytes)
        return True
    except Exception:
        return False

# function to create directory
def create_directory(directory_path):
    Path(directory_path).mkdir(parents=True, exist_ok=True)
    return Path(directory_path).is_dir()

# function to check directory existence
def directory_exists(directory_path):
    return os.path.isdir(directory_path)

# function to clear screen
def clear_current_screen():
    print("\033[2J\033[H", end="", flush=True)
    # uses ANSI escape sequences

# function to get base64 of file contents
def file_to_base64(file_path):
    path = Path(file_path)
    if not path.is_file():
        return False
    with path.open("rb") as f:
        data = f.read()
    return base64.b64encode(data).decode("ascii")

# some send email functions
# these will further call functions from our "tokenlook.py" code implementation
def send_email_without_attachment(jwt_token_value, proxy_list) -> None:
    # Instantiate the Logging and saving class
    ProcessLogJSON = ProcessLogsAndJson(log_file, config_file)

    # Instantiate the TokenLook OutlookConnect class
    SendMailConnector = tokenlook.OutlookConnect(jwt_token_value, proxy_list)

    # Ask User Input to Start Crafting the Email message:
    # Craft email content
    print(f"{bred}===> Crafting email to send <==={noclr}\n")
    ProcessLogJSON.write_tool_log(f"{bred}===> Crafting email to send <==={noclr}")

    msg_subject = input(f"[{bblue}ASK{noclr}] Enter the email subject to send:-> ")
    content_type = input(f"[{bblue}ASK{noclr}] Enter the content type (input 'Text' or 'HTML'):-> ")

    if content_type != 'HTML':
        print(f"[{bblue}ASK{noclr}] Enter the content body (Input 'end-of-body' and press ENTER to finish):")

        lines = []
        while True:
            line = input()
            if line == "end-of-body":
                break
            lines.append(line)
        content_body = "\n".join(lines)

    elif content_type == 'HTML':
        content_type_html_path = input(f"[{bblue}ASK{noclr}] Enter the complete path of HTML file:-> ")
        content_body = get_html_file_content(content_type_html_path)
    else:
        print(f"\n[{bred}!{noclr}] Invalid content type or specified type unsupported!")
        return # exit function

    to_rcpt = input(f"[{bblue}ASK{noclr}] Enter the recipient email:-> ")
    save_to_sent = input(f"[{bblue}ASK{noclr}] Save email in user's sent folder? [True/False]:-> ")

    # send the email
    print(f"\n[{bblue}INFO{noclr}] GET {bgreen}'/api/v2.0/me/sendmail'{noclr}")

    rcv_resp = SendMailConnector.send_mail_on_fly_no_attach(msg_subject, content_type, content_body, to_rcpt,
                                                            save_to_sent)

    if rcv_resp.status_code == 202:
        print(f"[{bblue}INFO{noclr}] status code {bgreen}{rcv_resp.status_code} OK{noclr}")
    else:
        print(f"[{bblue}INFO{noclr}] status code {bred}{rcv_resp.status_code}{noclr}")
        print(f"[{bred}!{noclr}] something went wrong!")
        return


# Send email without attachment
def send_email_with_attachment(jwt_token_value, proxy_list) -> None:
    # Instantiate the Logging and saving class
    ProcessLogJSON = ProcessLogsAndJson(log_file, config_file)

    # Instantiate the TokenLook OutlookConnect class
    SendMailConnector = tokenlook.OutlookConnect(jwt_token_value, proxy_list)

    # Ask User Input to Start Crafting the Email message:
    # Craft email content
    print(f"{bred}===> Crafting email to send <==={noclr}\n")
    ProcessLogJSON.write_tool_log(f"{bred}===> Crafting email to send <==={noclr}")

    msg_subject = input(f"[{bblue}ASK{noclr}] Enter the email subject to send:-> ")
    content_type = input(f"[{bblue}ASK{noclr}] Enter the content type (input 'Text' or 'HTML'):-> ")

    if content_type != 'HTML':
        print(f"[{bblue}ASK{noclr}] Enter the content body (Input 'end-of-body' and press ENTER to finish):")

        lines = []
        while True:
            line = input()
            if line == "end-of-body":
                break
            lines.append(line)
        content_body = "\n".join(lines)

    elif content_type == 'HTML':
        content_type_html_path = input(f"[{bblue}ASK{noclr}] Enter the complete path of HTML file:-> ")
        content_body = get_html_file_content(content_type_html_path)
    else:
        print(f"\n[{bred}!{noclr}] Invalid content type or specified type unsupported!")
        return # exit function

    to_rcpt = input(f"[{bblue}ASK{noclr}] Enter the recipient email:-> ")
    attachment_file_path = input(f"[{bblue}ASK{noclr}] Enter the full path of attachment file:-> ")
    save_to_sent = input(f"[{bblue}ASK{noclr}] Save email in user's sent folder? [True/False]:-> ")
    custom_file_name = input(f"[{bblue}ASK{noclr}] Enter a custom filename if required (For same name as original, you can enter the same name as original file ):-> ")

    file_base64_content = file_to_base64(attachment_file_path)

    if file_base64_content:
        print(f"[{bblue}INFO{noclr}] file contents processed successfully")
    else:
        print(f"[{bred}ERROR{noclr}] unable to process attachment file contents")

    # make email request
    print(f"\n[{bblue}INFO{noclr}] GET {bgreen}'/api/v2.0/me/sendmail'{noclr}")

    rcv_resp = SendMailConnector.send_mail_on_fly_with_attach(msg_subject, content_type, content_body, to_rcpt, save_to_sent, file_base64_content, custom_file_name)

    if rcv_resp.status_code == 202:
        print(f"[{bblue}INFO{noclr}] status code {bgreen}{rcv_resp.status_code} OK{noclr}")
        print(f"[{bblue}INFO{noclr}] {bblue}Message ID:{noclr} {cyan}{rcv_resp.json().get("Id", "")}{noclr}")
    else:
        print(f"[{bblue}INFO{noclr}] status code {bred}{rcv_resp.status_code}{noclr}")
        print(f"[{bred}!{noclr}] something went wrong!")
        return

def reply_to_msg_single(jwt_token_value, proxy_list):
    # Instantiate the Logging and saving class
    ProcessLogJSON = ProcessLogsAndJson(log_file, config_file)

    # Instantiate the TokenLook OutlookConnect class
    SendMailConnector = tokenlook.OutlookConnect(jwt_token_value, proxy_list)

    # Ask User Input to Start Crafting the Email message:
    # Craft email content
    print(f"{bred}===> Crafting email reply to send (Reply only to addresses in From field) <==={noclr}\n")
    ProcessLogJSON.write_tool_log(f"{bred}===> Crafting email reply to send (Reply only to addresses in From field) <==={noclr}\n")
    msg_id = input(f"[{bblue}ASK{noclr}] Enter the message id of email to reply to (find it as value 'Id' from the mail JSON files):-> ")

    print(f"[{bblue}ASK{noclr}] Enter the email content to send in reply (Input 'end-of-body' and press ENTER to finish):")

    lines = []
    while True:
        line = input()
        if line == "end-of-body":
            break
        lines.append(line)
    content_body = "\n".join(lines)

    # make email request
    print(f"\n[{bblue}INFO{noclr}] GET {bgreen}'/api/v2.0/me/sendmail/reply'{noclr}")

    rcv_resp = SendMailConnector.reply_message_single(msg_id, content_body)
    if rcv_resp.status_code == 202:
        print(f"[{bblue}INFO{noclr}] status code {bgreen}{rcv_resp.status_code} OK{noclr}")
    else:
        print(f"[{bblue}INFO{noclr}] status code {bred}{rcv_resp.status_code}{noclr}")
        print(f"[{bred}ERROR{noclr}] something went wrong!")
        return


def reply_to_msg_all(jwt_token_value, proxy_list):
    # Instantiate the Logging and saving class
    ProcessLogJSON = ProcessLogsAndJson(log_file, config_file)

    # Instantiate the TokenLook OutlookConnect class
    SendMailConnector = tokenlook.OutlookConnect(jwt_token_value, proxy_list)

    # Ask User Input to Start Crafting the Email message:
    # Craft email content
    print(f"{bred}===> Crafting email reply to send (Reply to all emails in email thread) <==={noclr}\n")
    ProcessLogJSON.write_tool_log(f"{bred}===> Crafting email reply to send (Reply to all emails in email thread) <==={noclr}\n")
    msg_id = input(f"[{bblue}ASK{noclr}] Enter the message id of email to reply to (find it as value 'Id' from the mail JSON files):-> ")

    print(f"[{bblue}ASK{noclr}] Enter the email content to send in reply (Input 'end-of-body' and press ENTER to finish):")

    lines = []
    while True:
        line = input()
        if line == "end-of-body":
            break
        lines.append(line)
    content_body = "\n".join(lines)

    # make email request
    print(f"\n[{bblue}INFO{noclr}] GET {bgreen}'/api/v2.0/me/sendmail/replyall'{noclr}")

    rcv_resp = SendMailConnector.reply_message_single(msg_id, content_body)
    if rcv_resp.status_code == 202:
        print(f"[{bblue}INFO{noclr}] status code {bgreen}{rcv_resp.status_code} OK{noclr}")
    else:
        print(f"[{bblue}INFO{noclr}] status code {bred}{rcv_resp.status_code}{noclr}")
        print(f"[{bred}ERROR{noclr}] something went wrong!")
        return

def create_a_draft_msg_no_attach(jwt_token_value, proxy_list):
    # Instantiate the Logging and saving class
    ProcessLogJSON = ProcessLogsAndJson(log_file, config_file)

    # Instantiate the TokenLook OutlookConnect class
    SendMailConnector = tokenlook.OutlookConnect(jwt_token_value, proxy_list)

    # Ask User Input to Start Crafting the Email message:
    # Craft email content
    print(f"{bred}===> Crafting DRAFT email (no attachment) <==={noclr}\n")
    ProcessLogJSON.write_tool_log(f"{bred}===> Crafting email to send <==={noclr}")

    msg_subject = input(f"[{bblue}ASK{noclr}] Enter the email subject to send:-> ")
    content_type = input(f"[{bblue}ASK{noclr}] Enter the content type (input 'Text' or 'HTML'):-> ")

    if content_type != 'HTML':
        print(f"[{bblue}ASK{noclr}] Enter the content body (Input 'end-of-body' and press ENTER to finish):")

        lines = []
        while True:
            line = input()
            if line == "end-of-body":
                break
            lines.append(line)
        content_body = "\n".join(lines)

    elif content_type == 'HTML':
        content_type_html_path = input(f"[{bblue}ASK{noclr}] Enter the complete path of HTML file:-> ")
        content_body = get_html_file_content(content_type_html_path)
    else:
        print(f"\n[{bred}!{noclr}] Invalid content type or specified type unsupported!")
        return # exit function

    to_rcpt = input(f"[{bblue}ASK{noclr}] Enter the recipient email:-> ")

    # send the email
    print(f"\n[{bblue}INFO{noclr}] GET {bgreen}'/api/v2.0/me/messages'{noclr}")

    rcv_resp = SendMailConnector.create_draft_message_no_attach(msg_subject, content_type, content_body, to_rcpt)

    if rcv_resp.status_code == 201:
        print(f"[{bblue}INFO{noclr}] status code {bgreen}{rcv_resp.status_code} OK{noclr}")
        print(f'{bred}{bblue}[{noclr}DRAFT message ID{bblue}]{noclr}{noclr}> {rcv_resp.json().get("Id", "")}')
    else:
        print(f"[{bblue}INFO{noclr}] status code {bred}{rcv_resp.status_code}{noclr}")
        print(f"[{bred}!{noclr}] something went wrong!")
        return

def send_draft_msg(jwt_token_value, proxy_list):
    # Instantiate the Logging and saving class
    ProcessLogJSON = ProcessLogsAndJson(log_file, config_file)

    # Instantiate the TokenLook OutlookConnect class
    SendMailConnector = tokenlook.OutlookConnect(jwt_token_value, proxy_list)

    # Ask User Input to Start Crafting the Email message:
    # Craft email content
    print(f"{bred}===> Sending a DRAFT message <==={noclr}\n")
    ProcessLogJSON.write_tool_log(f"{bred}===> Sending a DRAFT message <==={noclr}\n")
    msg_id = input(f"[{bblue}ASK{noclr}] Enter the message id of DRAFT email:-> ")

    # make email request
    print(f"\n[{bblue}INFO{noclr}] GET {bgreen}'/api/v2.0/me/sendmail/<message_id>/send'{noclr}")

    rcv_resp = SendMailConnector.send_draft_message(msg_id)
    if rcv_resp.status_code == 202:
        print(f"[{bblue}INFO{noclr}] status code {bgreen}{rcv_resp.status_code} OK{noclr}")
    else:
        print(f"[{bblue}INFO{noclr}] status code {bred}{rcv_resp.status_code}{noclr}")
        print(f"[{bred}ERROR{noclr}] something went wrong!")
        return

def delete_mail_msg(jwt_token_value, proxy_list):
    # Instantiate the Logging and saving class
    ProcessLogJSON = ProcessLogsAndJson(log_file, config_file)

    # Instantiate the TokenLook OutlookConnect class
    SendMailConnector = tokenlook.OutlookConnect(jwt_token_value, proxy_list)

    # Ask User Input to Start Crafting the Email message:
    # Craft email content
    print(f"{bred}===> Delete an email message (Permanently deleted) <==={noclr}\n")
    ProcessLogJSON.write_tool_log(f"{bred}===> Delete an email message (Permanently deleted) <==={noclr}\n")
    msg_id = input(f"[{bblue}ASK{noclr}] Enter the message id of email to delete (specify multiple IDs separated by comma to delete multiple emails):-> ")

    try:
        msg_ids = msg_id.split(",")
        for msg_id_item in msg_ids:
            print(f"\n[{bblue}INFO{noclr}] DELETE {bgreen}'/api/v2.0/me/messages/<message_id>'{noclr}")

            rcv_resp = SendMailConnector.delete_mail_message(msg_id_item)
            if rcv_resp.status_code == 204:
                print(f"[{bblue}INFO{noclr}] status code {bgreen}{rcv_resp.status_code} OK{noclr}")
            else:
                print(f"[{bblue}INFO{noclr}] status code {bred}{rcv_resp.status_code}{noclr}")
                print(f"[{bred}ERROR{noclr}] something went wrong!")
                return
    except Exception as e:
        print(f"\n[{bblue}INFO{noclr}] DELETE {bgreen}'/api/v2.0/me/messages/<message_id>'{noclr}")

        rcv_resp = SendMailConnector.delete_mail_message(msg_id)
        if rcv_resp.status_code == 204:
            print(f"[{bblue}INFO{noclr}] status code {bgreen}{rcv_resp.status_code} OK{noclr}")
        else:
            print(f"[{bblue}INFO{noclr}] status code {bred}{rcv_resp.status_code}{noclr}")
            print(f"[{bred}ERROR{noclr}] something went wrong!")
            return

def get_current_autoreply(jwt_token_value, proxy_list):
    # Instantiate the Logging and saving class
    ProcessLogJSON = ProcessLogsAndJson(log_file, config_file)

    # Instantiate the TokenLook OutlookConnect class
    SendMailConnector = tokenlook.OutlookConnect(jwt_token_value, proxy_list)

    print(f"\n[{bblue}INFO{noclr}] GET {bgreen}'/api/v2.0/me/MailboxSettings/AutomaticRepliesSetting'{noclr}")


    rcv_resp = SendMailConnector.get_auto_reply_settings()
    if rcv_resp.status_code == 200:
        print(f"[{bblue}INFO{noclr}] status code {bgreen}{rcv_resp.status_code} OK{noclr}")
    else:
        print(f"[{bblue}INFO{noclr}] status code {bred}{rcv_resp.status_code}{noclr}")
        print(f"[{bred}ERROR{noclr}] something went wrong!")
        return

    print(f"\n{bgreen}===> Current AutoReply settings are displayed below <==={noclr}\n")
    str_long = f"""
    {bblue}Status:{noclr} {orng}{rcv_resp.json().get("Status", "")}{noclr}
    {bblue}Starts:{noclr} {orng}{rcv_resp.json().get("ScheduledStartDateTime", "").get("DateTime", "")}{noclr} :: {bgreen}Timezone:{noclr} {rcv_resp.json().get("ScheduledStartDateTime", "").get("TimeZone", "")}
    {bblue}Ends:{noclr} {orng}{rcv_resp.json().get("ScheduledEndDateTime", "").get("DateTime", "")}{noclr} :: {bgreen}Timezone:{noclr} {rcv_resp.json().get("ScheduledEndDateTime", "").get("TimeZone", "")}
    {bblue}Internal reply message:{noclr} {orng}{rcv_resp.json().get("InternalReplyMessage", "")}{noclr}
    {bblue}External reply message:{noclr} {orng}{rcv_resp.json().get("ExternalReplyMessage", "")}{noclr}
    {bblue}ExternalAuidence:{noclr} {orng}{rcv_resp.json().get("ExternalAudience", "")}{noclr}
    """
    print(str_long)

def update_autoreply(jwt_token_value, proxy_list):
    # Instantiate the Logging and saving class
    ProcessLogJSON = ProcessLogsAndJson(log_file, config_file)

    # Instantiate the TokenLook OutlookConnect class
    SendMailConnector = tokenlook.OutlookConnect(jwt_token_value, proxy_list)

    # Ask User Input to Start Crafting the Email message:
    # Craft email content
    print(f"{bred}===> Crafting auto reply email <==={noclr}\n")
    ProcessLogJSON.write_tool_log(f"{bred}===> Crafting auto reply email <==={noclr}")

    start_time = input(f"[{bblue}ASK{noclr}] Enter the start time in YYYY-MM-DD-HH(hours)-MM(minutes) - {cyan}[Enter UTC in 24 hours format for hours and minutes]{noclr}:-> ")
    end_time = input(
        f"[{bblue}ASK{noclr}] Enter the end time in YYYY-MM-DD-HH(hours)-MM(minutes) - {cyan}[Enter UTC in 24 hours format for hours and minutes]{noclr}:-> ")

    print(f"[{bblue}ASK{noclr}] Enter the content body (Input 'end-of-body' and press ENTER to finish, note that same reply will be INT and EXT):")
    lines = []
    while True:
        line = input()
        if line == "end-of-body":
            break
        lines.append(line)
    content_body = "\n".join(lines)

    print(f"\n[{bblue}INFO{noclr}] GET {bgreen}'/api/v2.0/me/MailboxSettings'{noclr}")
    rcv_resp = SendMailConnector.update_auto_reply_settings(start_time, end_time, content_body, content_body)
    if rcv_resp.status_code == 200:
        print(f"[{bblue}INFO{noclr}] status code {bgreen}{rcv_resp.status_code} OK{noclr}")
    else:
        print(f"[{bblue}INFO{noclr}] status code {bred}{rcv_resp.status_code}{noclr}")
        print(f"[{bred}ERROR{noclr}] something went wrong!")
        return


def process_attachments(jwt_token_value, proxy_list, data_dir_attachments):
    # Instantiate the Logging and saving class
    ProcessLogJSON = ProcessLogsAndJson(log_file, config_file)

    # Instantiate the TokenLook OutlookConnect class
    SendMailConnector = tokenlook.OutlookConnect(jwt_token_value, proxy_list)

    if directory_exists(data_dir_attachments):
        print(f"\n[{bblue}INFO{noclr}] {bred}On-disk attachments will be saved to:{noclr} {cyan}{user_data_dir}{noclr}")
    else:
        print(f"[{bred}ERROR{noclr}] User directory: {cyan}{user_data_dir}{noclr} not found")
        return

    print(f"\n{bgreen}===> Email attachment processor <==={noclr}\n")

    mail_id = input(f"[{bblue}ASK{noclr}] Enter the msg-id of the email to process attachments for (For multiple separate by space):-> ")

    try:
        lst_mail_id = tuple(mail_id.split(""))
    except Exception as e:
        lst_mail_id = tuple([mail_id])

    for mail_id in lst_mail_id:
        print(f"\n[{bblue}INFO{noclr}] Processing: {bgreen}{mail_id}{noclr}")
        print(f"[{bblue}INFO{noclr}] Retrieving details for attachment(s)")
        print(f"[{bblue}INFO{noclr}] GET {bgreen}'/api/v2.0/me/MailboxSettings'{noclr}")
        rcv_get_attachments = SendMailConnector.get_attachment_list(mail_id)
        if rcv_get_attachments.status_code == 200:
            print(f"[{bblue}INFO{noclr}] status code {bgreen}{rcv_get_attachments.status_code} OK{noclr}")
            attach_values = rcv_get_attachments.json().get("value", "")
            if attach_values:
                for item in attach_values:
                    print(f"\n[{bgreen
                    }+{noclr}] {bblue}name:{noclr} {cyan}{item.get('Name', '')}{noclr} {bblue}Id:{noclr} {cyan}{item.get('Id', '')}{noclr}")
            else:
                print(f"\n[{bblue}INFO{noclr}] no attachments found! exiting current option")
                return
        else:
            print(f"[{bblue}INFO{noclr}] status code {bred}{rcv_get_attachments.status_code}{noclr}")
            print(f"[{bred}ERROR{noclr}] something went wrong!")
            return

        # start downloading attachments
        attach_ids = input(
            f"\n[{bblue}ASK{noclr}] Enter the ID of the attachment to download (For multiple separate by space):-> ")
        try:
            lst_attach_ids = tuple(attach_ids.split(" "))
        except Exception as e:
            lst_attach_ids = tuple([attach_ids])

        for item in lst_attach_ids:
            rcv_download_attachments = SendMailConnector.get_attachment_contents(mail_id, item)
            if rcv_download_attachments.status_code == 200:
                print(f"\n[{bblue}INFO{noclr}] status code {bgreen}{rcv_download_attachments.status_code} OK{noclr}")
                attach_name = rcv_download_attachments.json().get("Name", "")
                attach_contents = rcv_download_attachments.json().get("ContentBytes", "")
                print(f"[{bblue}INFO{noclr}] saving attachment {cyan}{attach_name}{noclr}")
                if save_file_bytes(attach_name, data_dir_attachments, attach_contents):
                    print(f"[{bgreen}+{noclr}] attachment: {cyan}{attach_name}{noclr} is saved to disk")
                else:
                    print(f"[{bred}ERROR{noclr}] something went wrong!, unable to save attachment")
                    pass
            else:
                print(f"[{bblue}INFO{noclr}] status code {bred}{rcv_download_attachments.status_code}{noclr}")
                print(f"[{bred}ERROR{noclr}] something went wrong!")
                return

def delete_attachments(jwt_token_value, proxy_list):
    # Instantiate the Logging and saving class
    ProcessLogJSON = ProcessLogsAndJson(log_file, config_file)

    # Instantiate the TokenLook OutlookConnect class
    SendMailConnector = tokenlook.OutlookConnect(jwt_token_value, proxy_list)


    print(f"\n{bgreen}===> Email attachment delete <==={noclr}\n")

    mail_id = input(f"[{bblue}ASK{noclr}] Enter the msg-id of the email to process attachments for (For multiple separate by space):-> ")

    try:
        lst_mail_id = tuple(mail_id.split(""))
    except Exception as e:
        lst_mail_id = tuple([mail_id])

    for mail_id in lst_mail_id:
        print(f"\n[{bblue}INFO{noclr}] Processing: {bgreen}{mail_id}{noclr}")
        print(f"[{bblue}INFO{noclr}] Retrieving details for attachment(s)")
        print(f"[{bblue}INFO{noclr}] GET {bgreen}'/api/v2.0/me/MailboxSettings'{noclr}")
        rcv_get_attachments = SendMailConnector.get_attachment_list(mail_id)
        if rcv_get_attachments.status_code == 200:
            print(f"[{bblue}INFO{noclr}] status code {bgreen}{rcv_get_attachments.status_code} OK{noclr}")
            attach_values = rcv_get_attachments.json().get("value", "")
            if attach_values:
                for item in attach_values:
                    print(f"\n[{bgreen
                    }+{noclr}] {bblue}name:{noclr} {cyan}{item.get('Name', '')}{noclr} {bblue}Id:{noclr} {cyan}{item.get('Id', '')}{noclr}")
            else:
                print(f"\n[{bblue}INFO{noclr}] no attachments found! exiting current option")
                return
        else:
            print(f"[{bblue}INFO{noclr}] status code {bred}{rcv_get_attachments.status_code}{noclr}")
            print(f"[{bred}ERROR{noclr}] something went wrong!")
            return

        # start deleting
        attach_ids = input(
            f"\n[{bblue}ASK{noclr}] Enter the ID of the attachment to delete (For multiple separate by space):-> ")
        try:
            lst_attach_ids = tuple(attach_ids.split(" "))
        except Exception as e:
            lst_attach_ids = tuple([attach_ids])

        for item in lst_attach_ids:
            rcv_del_attachments = SendMailConnector.delete_mail_attachment(mail_id, item)
            if rcv_del_attachments.status_code == 200 or 204:
                print(f"\n[{bblue}INFO{noclr}] status code {bgreen}{rcv_del_attachments.status_code} OK{noclr}")
                print(f"[{bgreen}+{noclr}] attachment {bred}deleted!{noclr}")
            else:
                print(f"[{bblue}INFO{noclr}] status code {bred}{rcv_del_attachments.status_code}{noclr}")
                print(f"[{bred}ERROR{noclr}] something went wrong!")
                return


if __name__ == "__main__":
    # Instantiate the class
    PLogJSON = ProcessLogsAndJson(log_file, config_file)
    print(f"\n[{bgreen}+{noclr}] Initiating TokenLook {bgreen}v{VERSION_MAIN}{noclr} execution at {bblue}{get_current_timestamp()}{noclr}")
    PLogJSON.write_tool_log(f"\n[{bgreen}+{noclr}] Initiating token look execution at {bblue}{get_current_timestamp()}{noclr}")

    # Check config file existence
    if PLogJSON.file_exists():
        PLogJSON.write_tool_log(f"[{bgreen}+{noclr}] config file found")
        print(f"[{bgreen}+{noclr}] config file found")
    else:
        PLogJSON.write_tool_log(f"[{bred}!{noclr}] config file not found! refer to instructions in README.md")
        print(f"[{bred}ERROR{noclr}] config file not found! refer to instructions in README.md")
        sys.exit()

    # Parse config
    config_values = PLogJSON.return_json_data()
    # Process values
    jwt_token_list = config_values.get("current_jwt", "")
    data_dir = config_values.get("data_dir", "")
    net_proxy = config_values.get("proxy", "")
    mail_numb = config_values.get("mail_numb", "")

    if not jwt_token_list or not data_dir or not net_proxy:
        print(f"[{bred}ERROR{noclr}] config file missing a value! refer to instructions in README.md")
        PLogJSON.write_tool_log(f"[{bred}ERROR{noclr}] config file missing a value! refer to instructions in README.md")


    for jwt_token in jwt_token_list:
        # check existence for data directory and create it if not present
        if directory_exists(data_dir):
            print(f"[{bgreen}+{noclr}] Data directory found")
            PLogJSON.write_tool_log(f"[{bgreen}+{noclr}] Data directory found")
        else:
            print(f"[{bred}ERROR{noclr}] Data directory not found!")
            PLogJSON.write_tool_log(f"[{bred}ERROR{noclr}] Data directory not found!")
            if create_directory(data_dir):
                print(f"[{bgreen}+{noclr}] Data directory created {bblue}{data_dir}{noclr}")
                PLogJSON.write_tool_log(f"[{bgreen}+{noclr}] Data directory created {bblue}{data_dir}{noclr}")
            else:
                print(f"[{bred}ERROR{noclr}] unable to create data directory!")
                PLogJSON.write_tool_log(f"[{bred}ERROR{noclr}] unable to create data directory!")

        # Finally parse JWT token segments
        header, payload, signature = jwt_token.split(".")

        # Instantiate JWT process class
        jwtProcess = tokenlook.ParseMyJwt(header, payload, signature)

        # Process header segment
        jwt_header = jwtProcess.parse_header()
        PLogJSON.write_tool_log(f"\n[{bgreen}+{noclr}] Parsing JWT header")
        print(f"\n[{bgreen}+{noclr}] Parsing JWT header")
        header_prt = f"""
        {bblue}typ:{noclr} {bgreen}{jwt_header.get("typ", "")}{noclr}
        {bblue}alg:{noclr} {bgreen}{jwt_header.get("alg", "")}{noclr}
        {bblue}nonce:{noclr} {bgreen}{jwt_header.get("nonce", "")}{noclr}
        """
        print(header_prt)

        # Process Payload segment
        jwt_payload = jwtProcess.parse_payload()

        PLogJSON.write_tool_log(f"\n[{bgreen}+{noclr}] Parsing JWT payload")
        print(f"\n[{bgreen}+{noclr}] Parsing JWT payload")
        payload_prt = f"""
        {bred}aud:{noclr} {cyan}{jwt_payload.get("aud", "")}{noclr}
        {bred}exp_time:{noclr} {cyan}{datetime.fromtimestamp(int(jwt_payload.get("exp", "")))}{noclr}
        {bblue}name:{noclr} {cyan}{jwt_payload.get("name", "")}{noclr}
        {bblue}upn:{noclr} {cyan}{jwt_payload.get("upn", "")}{noclr}
        {bblue}app_displayname:{noclr} {cyan}{jwt_payload.get("app_displayname", "")}{noclr}
        """
        print(payload_prt)

        # Process signature
        print(f"\n[{bgreen}+{noclr}] Partially verifying JWT signature")
        PLogJSON.write_tool_log(f"\n[{bgreen}+{noclr}] Partially verifying JWT signature")
        if jwtProcess.parse_verify_signature():
            print(f"[{bgreen}+{noclr}] signature length {orng}OK{noclr}")
            PLogJSON.write_tool_log(f"[{bgreen}+{noclr}] signature length {orng}OK{noclr}")
        else:
            print(f"[{bred}ERROR{noclr}] signature length incorrect, JWT might be invalid")
            PLogJSON.write_tool_log(f"[{bred}ERROR{noclr}] signature length incorrect, JWT might be invalid")

        # Instantiate Outlook Connect
        OutlookConnector = tokenlook.OutlookConnect(jwt_token, net_proxy)

        # Get logs
        print(f"\n[{bblue}INFO{noclr}] using proxie(s) {bgreen}{net_proxy}{noclr}")
        PLogJSON.write_tool_log(f"\n[{bblue}INFO{noclr}] using proxie(s) {bgreen}{net_proxy}{noclr}")


        # Send request to me endpoint
        print(f"\n[{bblue}INFO{noclr}] GET {bgreen}'/api/v2.0/me'{noclr}")
        PLogJSON.write_tool_log(f"\n[{bblue}INFO{noclr}] GET {bgreen}'/api/v2.0/me'{noclr}")
        send_me_req = OutlookConnector.get_profile()

        # status code
        if send_me_req.status_code == 200:
            PLogJSON.write_tool_log(f"[{bblue}INFO{noclr}] status code {bgreen}{send_me_req.status_code} OK{noclr}")
            print(f"[{bblue}INFO{noclr}] status code {bgreen}{send_me_req.status_code} OK{noclr}")
            # Get JSON
            resp_json = send_me_req.json()
        else:
            print(f"[{bblue}INFO{noclr}] status code {bred}{send_me_req.status_code}{noclr}")
            sys.exit(0)
        PLogJSON.write_tool_log(f"[{bgreen}+{noclr}] response data")
        print(f"[{bgreen}+{noclr}] response data")
        resp_json_str = f"""
            {bblue}Email:{noclr} {cyan}{resp_json.get("EmailAddress", "")}{noclr}
            {bblue}DisplayName:{noclr} {cyan}{resp_json.get("DisplayName", "")}{noclr}
            {bblue}MailboxGuid:{noclr} {cyan}{resp_json.get("MailboxGuid", "")}{noclr}
        """
        print(resp_json_str)

        # Start getting mail messages multiple
        # we will create a local disk copy of all the emails.
        # we will create another directory inside the data directory specific to the user
        user_data_dir = data_dir + "/" + resp_json.get("DisplayName").split(" ")[0] + "_"  + resp_json.get("MailboxGuid")
        print(f"[{bblue}INFO{noclr}] extracted emails will be saved to {cyan}{user_data_dir}{noclr}")
        PLogJSON.write_tool_log(f"[{bblue}INFO{noclr}] extracted emails will be saved to {cyan}{user_data_dir}{noclr}")

        if not directory_exists(user_data_dir):
            print(f"[{bblue}INFO{noclr}] creating user directory {cyan}{user_data_dir}{noclr}")
            PLogJSON.write_tool_log(f"[{bblue}INFO{noclr}] creating user directory {cyan}{user_data_dir}{noclr}")
            create_directory(user_data_dir)
        else:
            print(f"\n[{bgreen}+{noclr}] user directory {cyan}{user_data_dir}{noclr} already exists")
            PLogJSON.write_tool_log(f"\n[{bgreen}+{noclr}] user directory {cyan}{user_data_dir}{noclr} already exists")

        # Start the loop to extract and save emails to disk, 10 emails in each request
        skip_numb = 0
        top_numb = 10

        while skip_numb < mail_numb:
            emails_rsp = OutlookConnector.get_mail_messages_multiple(skip_numb, top_numb) # JSON response with emails
            print(f"\n[{bblue}INFO{noclr}] GET {bgreen}'/api/v2.0/me/messages'{noclr}")
            PLogJSON.write_tool_log(f"\n[{bblue}INFO{noclr}] GET {bgreen}'/api/v2.0/me/messages'{noclr}")

            if emails_rsp.status_code == 200:
                print(f"[{bblue}INFO{noclr}] status code {bgreen}{emails_rsp.status_code} OK{noclr}")
                PLogJSON.write_tool_log(f"[{bblue}INFO{noclr}] status code {bgreen}{emails_rsp.status_code} OK{noclr}")
            else:
                print(f"[{bblue}INFO{noclr}] status code {bred}{emails_rsp.status_code}{noclr}")
                PLogJSON.write_tool_log(f"[{bblue}INFO{noclr}] status code {bred}{emails_rsp.status_code}{noclr}")
                sys.exit(0)


            # loop through all emails to save each individual with message ID
            print(f"\n[{bblue}INFO{noclr}] emails will be saved to disk in the user directory")
            email_items_list = emails_rsp.json().get("value", "")
            for email_item in email_items_list:
                try:
                    time.sleep(0.5)
                    print(f"\n[{bgreen}+{noclr}] Processing email with {bblue}subject:{noclr} {orng}{email_item.get("Subject")}{noclr}")
                    PLogJSON.write_tool_log(f"\n[{bgreen}+{noclr}] Processing email with {bblue}subject:{noclr} {orng}{email_item.get("Subject")}{noclr}")

                    # From email
                    print(f"[{bgreen}+{noclr}] {bblue}From:{noclr} {cyan}{email_item.get("From", "").get("EmailAddress", "").get("Address", "")}{noclr}")
                    PLogJSON.write_tool_log(f"[{bgreen}+{noclr}] {bblue}From:{noclr} {cyan}{email_item.get("From", "").get("EmailAddress", "").get("Address", "")}{noclr}")

                    # To recipient
                    print(f"[{bgreen}+{noclr}] {bblue}ToRecipients-0:{noclr} {cyan}{email_item.get("ToRecipients", "")[0].get("EmailAddress", "").get("Address", "")}{noclr}")
                    PLogJSON.write_tool_log(f"[{bgreen}+{noclr}] {bblue}ToRecipients-0:{noclr} {cyan}{email_item.get("ToRecipients", "")[0].get("EmailAddress", "").get("Address", "")}{noclr}")

                    # file path to save to
                    disk_mail_path = user_data_dir + "/" + email_item.get("Id") + ".json"

                    # Get message JSON
                    mail_json = OutlookConnector.get_single_mail(email_item.get("Id"))
                    mail_json_resp = mail_json.json()
                    print(f"[{bblue}INFO{noclr}] GET {bgreen}'/api/v2.0/me/messages/<message-id>'{noclr}")
                    if mail_json.status_code == 200:
                        print(f"[{bblue}INFO{noclr}] status code {bgreen}{mail_json.status_code} OK{noclr}")
                        PLogJSON.write_tool_log(f"[{bblue}INFO{noclr}] status code {bgreen}{mail_json.status_code} OK{noclr}")
                    else:
                        print(f"[{bblue}INFO{noclr}] status code {bred}{mail_json_resp.status_code}{noclr}")
                        PLogJSON.write_tool_log(f"[{bblue}INFO{noclr}] status code {bred}{mail_json_resp.status_code}{noclr}")
                        sys.exit(0)

                    # Save mail on disk
                    save_json_resp(disk_mail_path, mail_json_resp)
                    print(f"[{bgreen}+{noclr}] {bblue}saved to disk:{noclr} {cyan}{disk_mail_path}{noclr}")
                    PLogJSON.write_tool_log(f"[{bgreen}+{noclr}] {bblue}saved to disk:{noclr} {cyan}{disk_mail_path}{noclr}")

                    # Process Mail HTML page as well
                    # Check if mail content type is HTML, if it is , save the content to disk in HTML file prefixed by message ID and user's first name
                    if "HTML" in mail_json_resp.get("Body", "").get("ContentType", ""):
                        disk_html_mail_path = user_data_dir + "/" + email_item.get("Id") + ".html"
                        print(f"[{bblue}INFO{noclr}] Mail content type is {bgreen}HTML like{noclr}, {orng}HTML content will be saved to disk{noclr}")
                        PLogJSON.write_tool_log(f"[{bblue}INFO{noclr}] Mail content type is {bgreen}HTML like{noclr}, {orng}HTML content will be saved to disk{noclr}")
                        email_html_content = mail_json_resp.get("Body", "").get("Content", "")
                        if save_mail_html(disk_html_mail_path, email_html_content):
                            print(f"[{bgreen}+{noclr}] HTML content is {cyan}saved to disk{noclr}")
                            PLogJSON.write_tool_log(f"[{bgreen}+{noclr}] HTML content is {cyan}saved to disk{noclr}")
                        else:
                            print(f"[{bred}ERROR{noclr}] unable to save HTML")
                            PLogJSON.write_tool_log(f"[{bred}!{noclr}] unable to save HTML")
                            pass
                except Exception as e:
                    print(f"[{bred}ERROR{noclr}] unable to fully process the mail, email likely a DRAFT message")
                    print(f"[{bblue}INFO{noclr}] Mail ID for reference {cyan}{email_item.get("Id")}{noclr}")
                    pass
            # increment the number
            skip_numb += 10

        print(f"\n[{bgreen}+{noclr}] Maximum specified emails {cyan}<={noclr}{bred}[{noclr}{bblue}{mail_numb}{noclr}{bred}]{noclr} have been extracted!")

        # Print current user context once again
        current_user_context = f"""\n
        {bred}===> Current User Context <==={noclr}
        {bblue}name:{noclr} {cyan}{jwt_payload.get("name", "")}{noclr}
        {bblue}upn:{noclr} {cyan}{jwt_payload.get("upn", "")}{noclr}
        {bred}exp_time:{noclr} {cyan}{datetime.fromtimestamp(int(jwt_payload.get("exp", "")))}{noclr}
        """

        print(current_user_context)

        # Run Functions and Code to send emails now.
        ask_line = f"""
[{bblue}ASK{noclr}] Initializing {bred}Send-email Menu for current user Context{noclr} Enter {bgreen}Y/y{noclr} to continue OR  {bred}N/n{noclr} to stop 
[{bblue}INFO{noclr}]{cyan} Note that selecting Y/y will automatically clear the console!{noclr}-->
"""
        ask_menu_start = input(ask_line)

        if ask_menu_start == "Y" or ask_menu_start == "y":
            menu(send_email_banner,
                 {
                     "Send Email without attachment": lambda: send_email_without_attachment(jwt_token, net_proxy),
                     "Send Email with attachment": lambda: send_email_with_attachment(jwt_token, net_proxy),
                     "Reply to a message (standard single reply)": lambda: reply_to_msg_single(jwt_token, net_proxy),
                     "Reply all to a message (reply all (same as mail app))": lambda: reply_to_msg_all(jwt_token, net_proxy),
                     "Create a draft message without attachment (saved to the Drafts folder)": lambda: create_a_draft_msg_no_attach(jwt_token, net_proxy),
                     "Send a draft message": lambda: send_draft_msg(jwt_token, net_proxy),
                     "Delete mail messages (Permanently Deleted)": lambda: delete_mail_msg(jwt_token, net_proxy),
                     "Get current auto reply settings (view current auto reply settings)": lambda : get_current_autoreply(jwt_token, net_proxy),
                     "Schedule auto reply [e.g OOO]:": lambda: update_autoreply(jwt_token, net_proxy),
                     "Get email attachment(s)": lambda: process_attachments(jwt_token, net_proxy, user_data_dir),
                     "Delete email attachment(s) (Permanently Deleted)": lambda: delete_attachments(jwt_token, net_proxy),
                     "Get folder collection (Get Info on all the Folders in mail client):": lambda: clear_current_screen(), # To be implemented
                     "Create new folder": lambda: clear_current_screen(), # To be implemented
                     "Update folder": lambda: clear_current_screen(), # To be implemented
                     "Delete folder": lambda: clear_current_screen(), # To be implemented
                     "Move or Copy folder": lambda: clear_current_screen(), # To be implemented
                     "Clear console (selecting exit will move to next menu)": lambda: clear_current_screen() # To be implemented
                 }
                 )

        elif ask_menu_start == "N" or ask_menu_start == "n":
            print(f"[{bred}ERROR{noclr}] Send-email Menu initialization for current user context {bred}cancelled!{noclr}, {cyan}moving to next user context!{noclr}")
            pass

        else:
            print(f"[{bred}ERROR{noclr}] Send-email Menu initialization for current user context {bred}failed!{noclr}, {cyan}moving to next user context!{noclr}")
            pass

        # Run Functions and Code to Initiate the Mail-Search now
        ask_line = f"""
        [{bblue}ASK{noclr}] Initializing {bred}Send-email Menu for current user Context{noclr} Enter {bgreen}Y/y{noclr} to continue OR  {bred}N/n{noclr} to stop 
        [{bblue}INFO{noclr}]{cyan} Note that selecting Y/y will automatically clear the console!{noclr}-->
        """
