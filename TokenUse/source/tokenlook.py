import random
import requests
import base64
import json
from datetime import datetime
import urllib3
import urllib.parse
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# For colored text in terminal output
bgreen = '\033[1;32m'
bred = '\033[1;31m'
noclr = '\033[0m'
bblue = '\033[1;34m'
cyan = '\033[0;36m'
orng = '\033[0;33m'


# time conversion
def to_dotnet_utc_format(user_input: str) -> str:
    # Parse input: YYYY-MM-DD-HH-MM
    dt = datetime.strptime(user_input, "%Y-%m-%d-%H-%M")

    # Format to .NET-style UTC string with 7-digit fractional seconds
    return str(dt.strftime("%Y-%m-%dT%H:%M:%S.0000000"))


# class with various methods to parse JWT
class ParseMyJwt:
    def __init__(self, jwt_header, payload, signature):
        self.jwt_header = jwt_header
        self.payload = payload
        self.signature = signature

    # Below are various methods to parse the JWT and extract meaningful info
    def parse_header(self):
        # JWT uses Base64URL without padding
        padding = '=' * (-len(self.jwt_header) % 4)
        decoded_bytes = base64.urlsafe_b64decode(self.jwt_header + padding)
        return json.loads(decoded_bytes)


    def parse_payload(self):
        # JWT uses Base64URL without padding
        padding = '=' * (-len(self.payload) % 4)
        decoded_bytes = base64.urlsafe_b64decode(self.payload + padding)
        return json.loads(decoded_bytes)

    def parse_verify_signature(self):
        length_signature = len(self.signature)
        if length_signature == 342:
            return True
        else:
            return False

class OutlookConnect:
    def __init__(self, jwt_token, proxy_list):
        self.jwt_token = jwt_token
        self.proxy_list = proxy_list

        self.headers = {
            "Prefer" : "outlook.allow-unsafe-html",
            "User-Agent" : "Microsoft Office/16.0 (Windows NT 10.0; Microsoft Outlook 16.0.19725; Pro)",
            "Accept-Encoding" : "gzip, deflate, br",
            "Accept": "application/json",
            "Authorization" : "Bearer " + self.jwt_token
        }

    def get_proxy_config(self):
        if not self.proxy_list:
            return None
        elif self.proxy_list == "":
            return None

        choose_proxy = random.choice(self.proxy_list)
        # make proxy config
        proxy_config = {
            "http": choose_proxy,
            "https": choose_proxy
        }
        return proxy_config


    # Simple me endpoint call to verify if token is valid
    def get_profile(self):
        proxy_list = self.get_proxy_config()

        endpoint_webpath = "https://outlook.office.com/api/v2.0/me"


        # use requests with session
        session = requests.Session()
        session.headers.update(self.headers)

        response = session.get(endpoint_webpath, proxies=proxy_list, timeout=30, verify=False)
        return response

    def get_mail_messages_multiple(self, skip_numb, top_numb):
        proxy_list = self.get_proxy_config()

        endpoint_webpath = f"https://outlook.office.com/api/v2.0/me/messages?skip={skip_numb}&top={top_numb}"

        # use requests with session
        session = requests.Session()
        session.headers.update(self.headers)

        response = session.get(endpoint_webpath, proxies=proxy_list, timeout=30, verify=False)
        return response

    def get_single_mail(self, message_id):
        proxy_list = self.get_proxy_config()

        endpoint_webpath = f"https://outlook.office.com/api/v2.0/me/messages/{message_id}"

        # use requests with session
        session = requests.Session()
        session.headers.update(self.headers)

        response = session.get(endpoint_webpath, proxies=proxy_list, timeout=30, verify=False)
        return response

    # Send a message on the fly
    def send_mail_on_fly_no_attach(self, msg_subject, content_type, content_body, to_rcpt, saveto_sent):
        """
        :param msg_subject: Subject of the email message
        :param content_type: Type of content to send in the email
        :param content_body: The content body to send in the message
        :param to_rcpt: Email of the receiving recipient
        :param saveto_sent: Bool value True or False to specify if file should be saved to Sent folder after its sent
        :return: will return the response of the request
        """
        proxy_list = self.get_proxy_config()

        endpoint_webpath = "https://outlook.office.com/api/v2.0/me/sendmail"

        # Build JSON data
        json_data = {
              "Message": {
                "Subject": msg_subject,
                "Body": {
                  "ContentType": content_type,
                  "Content": content_body
                },
                "ToRecipients": [
                  {
                    "EmailAddress": {
                      "Address": to_rcpt
                    }
                  }
                ]
              },
              "SaveToSentItems": saveto_sent
            }

        # use requests with session
        session = requests.Session()
        session.headers.update(self.headers)

        # add JSON content type header as well

        session.headers.update({
            "Content-Type": "application/json"
        })

        resp = session.post(
            endpoint_webpath,
            data=json.dumps(json_data),
            proxies=proxy_list,
            timeout=30,
            verify=False
        )

        return resp

    def send_mail_on_fly_with_attach(self, msg_subject, content_type, content_body, to_rcpt, saveto_sent, attach_file_bytes, custom_name):
        """
        :param msg_subject: Subject of the email message
        :param content_type: Type of content to send in the email
        :param content_body: The content body to send in the message
        :param to_rcpt: Email of the receiving recipient
        :param saveto_sent: Bool value True or False to specify if file should be saved to Sent folder after its sent
        :param attach_file_bytes: pass the variable storing the raw bytes of the file
        :param custom_name: Specify a custom name to the attached file, different from the original file name
        :return: returns response of the request
        """
        proxy_list = self.get_proxy_config()

        endpoint_webpath = "https://outlook.office.com/api/v2.0/me/sendmail"

        # Build JSON data
        json_data = {
            "Message": {
                "Subject": msg_subject,
                "Body": {
                    "ContentType": content_type,
                    "Content": content_body
                },
                "ToRecipients": [
                    {
                        "EmailAddress": {
                            "Address": to_rcpt
                        }
                    }
                ],
                "Attachments": [
                    {
                        "@odata.type": "#Microsoft.OutlookServices.FileAttachment",
                        "Name": custom_name,
                        "ContentBytes": attach_file_bytes
                    }
                ]
            },
            "SaveToSentItems": saveto_sent
        }

        # use requests with session
        session = requests.Session()
        session.headers.update(self.headers)

        # add JSON content type header as well

        session.headers.update({
            "Content-Type": "application/json"
        })

        resp = session.post(
            endpoint_webpath,
            data=json.dumps(json_data),
            proxies=proxy_list,
            timeout=30,
            verify=False
        )

        return resp


    def reply_message_single(self, msg_id, reply_content):
        """
        :param msg_id of the message to delete
        :param reply_content content of the message to reply
        :return: returns response of the request
        """
        proxy_list = self.get_proxy_config()

        endpoint_webpath = f"https://outlook.office.com/api/v2.0/me/messages/{msg_id}/reply"

        # Build JSON data
        json_data = {
            "Comment": reply_content
        }

        # use requests with session
        session = requests.Session()
        session.headers.update(self.headers)

        # add JSON content type header as well

        session.headers.update({
            "Content-Type": "application/json"
        })

        resp = session.post(
            endpoint_webpath,
            data=json.dumps(json_data),
            proxies=proxy_list,
            timeout=30,
            verify=False
        )

        return resp


    def reply_message_all(self, msg_id, reply_content):
        """
        :param msg_id of the message to delete
        :param reply_content content of the message to reply
        :return: returns response of the request
        """
        proxy_list = self.get_proxy_config()

        endpoint_webpath = f"https://outlook.office.com/api/v2.0/me/messages/{msg_id}/replyall"

        # Build JSON data
        json_data = {
            "Comment": reply_content
        }

        # use requests with session
        session = requests.Session()
        session.headers.update(self.headers)

        # add JSON content type header as well

        session.headers.update({
            "Content-Type": "application/json"
        })

        resp = session.post(
            endpoint_webpath,
            data=json.dumps(json_data),
            proxies=proxy_list,
            timeout=30,
            verify=False
        )

        return resp



    def create_draft_message_no_attach(self, msg_subject, content_type, content_body, to_rcpt,):
        """
        :param msg_subject: Subject of the email message
        :param content_type: Type of content to send in the email
        :param content_body: The content body to send in the message
        :param to_rcpt: Email of the receiving recipient
        :return: returns response of the request
        """
        proxy_list = self.get_proxy_config()

        endpoint_webpath = "https://outlook.office.com/api/v2.0/me/messages"

        # Build JSON data
        json_data = {
                "Subject": msg_subject,
                "Importance": "Low",
                "Body": {
                    "ContentType": content_type,
                    "Content": content_body
                },
                "ToRecipients": [
                    {
                        "EmailAddress": {
                            "Address": to_rcpt
                        }
                    }
                ]
            }

        # use requests with session
        session = requests.Session()
        session.headers.update(self.headers)

        # add JSON content type header as well
        session.headers.update({
            "Content-Type": "application/json"
        })

        resp = session.post(
            endpoint_webpath,
            data=json.dumps(json_data),
            proxies=proxy_list,
            timeout=30,
            verify=False
        )

        return resp

    def send_draft_message(self, msg_id):
        """
        :param msg_id: Message ID of the draft message, received in response when draft was created.
        :return: returns response of the request
        """
        proxy_list = self.get_proxy_config()

        endpoint_webpath = f"https://outlook.office.com/api/v2.0/me/messages/{msg_id}/send"

        # use requests with session
        session = requests.Session()
        session.headers.update(self.headers)

        resp = session.post(
            endpoint_webpath,
            proxies=proxy_list,
            timeout=30,
            verify=False
        )

        return resp

    def delete_mail_message(self, msg_id):
        """
        :param msg_id: Message ID of the draft message, received in response when draft was created.
        :return: returns response of the request
        """
        proxy_list = self.get_proxy_config()

        endpoint_webpath = f"https://outlook.office.com/api/v2.0/me/messages/{msg_id}"

        # use requests with session
        session = requests.Session()
        session.headers.update(self.headers)

        resp = session.delete(
            endpoint_webpath,
            proxies=proxy_list,
            timeout=30,
            verify=False
        )

        return resp


    def get_auto_reply_settings(self):
        """
        :return: returns response of the request
        """
        proxy_list = self.get_proxy_config()

        endpoint_webpath = "https://outlook.office.com/api/v2.0/me/MailboxSettings/AutomaticRepliesSetting"

        # use requests with session
        session = requests.Session()
        session.headers.update(self.headers)

        resp = session.get(
            endpoint_webpath,
            proxies=proxy_list,
            timeout=30,
            verify=False
        )

        return resp


    def update_auto_reply_settings(self, start_date, end_date, internal_reply, external_reply):
        """
        :return: returns response of the request
        """
        proxy_list = self.get_proxy_config()

        endpoint_webpath = "https://outlook.office.com/api/v2.0/me/MailboxSettings"

        # use requests with session
        session = requests.Session()
        session.headers.update(self.headers)

        # add JSON content type header as well
        session.headers.update({
            "Content-Type": "application/json"
        })

        json_data = {
            "@odata.context": "https://outlook.office.com/api/v2.0/$metadata#Me/MailboxSettings",
            "AutomaticRepliesSetting": {
                "Status": "Scheduled",
                "ScheduledStartDateTime": {
                  "DateTime": to_dotnet_utc_format(start_date),
                  "TimeZone": "UTC"
                },
                "ScheduledEndDateTime": {
                  "DateTime":  to_dotnet_utc_format(end_date),
                  "TimeZone": "UTC"
                },
                "InternalReplyMessage": internal_reply,
                "ExternalReplyMessage": external_reply
            }
        }

        resp = session.patch(
            endpoint_webpath,
            proxies=proxy_list,
            timeout=30,
            data=json.dumps(json_data),
            verify=False
        )

        return resp

    def get_attachment_list(self, msg_id):
        """
        :msg_id: Message ID of email to process for attachments.
        :return: returns response of the request
        """
        proxy_list = self.get_proxy_config()

        endpoint_webpath = f"https://outlook.office.com/api/v2.0/me/messages/{msg_id}/attachments?$select=Name"

        # use requests with session
        session = requests.Session()
        session.headers.update(self.headers)

        resp = session.get(
            endpoint_webpath,
            proxies=proxy_list,
            timeout=30,
            verify=False
        )
        return resp


    def get_attachment_contents(self, msg_id, attachment_id):
        """
        :msg_id: Message ID of email to process for attachments.
        :attachment_id: Attachment ID of email to process for attachments.
        :return: returns response of the request
        """
        proxy_list = self.get_proxy_config()

        safe_attachment_id = urllib.parse.quote(attachment_id, safe="")
        # this fixed the previous errors, not entirely sure why msg_id alone in previous requests was working fine but here some messed up encoding issue came in.

        endpoint_webpath = f"https://outlook.office.com/api/v2.0/me/messages/{msg_id}/attachments/{safe_attachment_id}"

        # use requests with session
        session = requests.Session()
        session.headers.update(self.headers)

        resp = session.get(
            endpoint_webpath,
            proxies=proxy_list,
            timeout=30,
            verify=False
        )
        return resp


    # deleted files are permanently gone, not even present in users person onedrive recyclebin
    def delete_mail_attachment(self, msg_id, attachment_id):
        """
        :param msg_id: Message ID of email to delete attachment(s).
        :param attachment_id: ID of the attachment to delete.
        :return: returns response of the request
        """

        proxy_list = self.get_proxy_config()

        safe_attachment_id = urllib.parse.quote(attachment_id, safe="")
        # this fixed the previous errors, not entirely sure why msg_id alone in previous requests was working fine but here some messed up encoding issue came in.

        endpoint_webpath = f"https://outlook.office.com/api/v2.0/me/messages/{msg_id}/attachments/{safe_attachment_id}"

        # use requests with session
        session = requests.Session()
        session.headers.update(self.headers)

        resp = session.delete(
            endpoint_webpath,
            proxies=proxy_list,
            timeout=30,
            verify=False
        )
        return resp
