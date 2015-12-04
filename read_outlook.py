from win32com.client.gencache import EnsureDispatch as Dispatch
import logging
import sys
import re

# set parameters
inbox_name = 'mymail@abc.eu'
folder_name = 'Inbox'
subfolder_name = ''
message_subject_filter = ''
email_regex = r'[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+'

outlook = Dispatch("Outlook.Application")
mapi = outlook.GetNamespace("MAPI")

logging.basicConfig(filename='outlook_read.log')
log = logging.getLogger()
log.setLevel(logging.DEBUG)

ch = logging.StreamHandler(sys.stdout)
ch.setLevel(logging.DEBUG)
formatter = logging.Formatter(
    '%(asctime)s - %(name)s - %(levelname)s - %(message)s')
ch.setFormatter(formatter)
log.addHandler(ch)

emails = []


class OutlookObj(object):
    def __init__(self, outlook_object):
        self._obj = outlook_object

    def items(self):
        array_size = self._obj.Count
        for item_index in xrange(1, array_size + 1):
            yield (item_index, self._obj[item_index])

    def prop(self):
        return sorted(self._obj._prop_map_get_.keys())


for inx_folder, folder in OutlookObj(mapi.Folders).items():
    # iterate all Outlook folders (top level)

    if folder.Name == inbox_name:
        # print "-" * 70
        log.info(folder.Name)

        for inx_subfolders, subfolder in OutlookObj(folder.Folders).items():
            if subfolder.Name == folder_name:
                log.info("(%i)" % inx_subfolders,
                         subfolder.Name, "=> ", subfolder)

                messages = subfolder.Items

                # Iterate through the messages contained within our subfolder
                for i, message in enumerate(messages, start=1):
                    # filter email with specific string in subject
                    if message_subject_filter in message.Subject:
                        try:
                            # Treat message as a singular object, you can use
                            # the body, sender, cc's and pretty much every
                            # aspect of an e-mail
                            # In my case I used body and subject
                            log.info('({}) {}'.format(i, message.Subject.encode(
                                'utf-8')))
                            
                            # get all emails from message's body
                            email = re.findall(email_regex, message.Body,
                                               re.MULTILINE)
                            log.info('Emails found: {}'.format(email))
                            emails.extend(email)
                        except Exception as err:
                            log.error(
                                "Error accessing mailItem: {}".format(str(err)))

# remove duplicated emails
emails = sorted(list(set(emails)))

log.info(emails)
with open('email.txt', 'w') as f:
    f.write("\n".join(emails))
