__author__  = "Witold Lawacz (wit0k)"
__date__    = "2019-01-15"
__version__ = '0.1.0.3'

import olefile as OleFile  # pip install olefile
import re
import argparse
import sys
import os.path
import glob
import platform as _os
from os.path import isfile, isdir
import iocextract

""" Set working directory so the script can be executed from any location/symlink """
os.chdir(os.path.dirname(os.path.abspath(__file__)))

CRED = '\033[91m'
CYELLOW = '\33[33m'
CEND = '\033[0m'

TYPE_RECIPIENT = 3

class Attachment:

    def __init__(self, msg, dir_):

        # Get long filename
        self.longFilename = msg._getStream([dir_, '__substg1.0_3707001F'])

        # Get short filename
        self.shortFilename = msg._getStream([dir_, '__substg1.0_3704001F'])

        # Get attachment data
        self.data = msg._getStream([dir_,'__substg1.0_37010102'])

    def get_unique_filename(self, file_path):

        _loop = True
        index = 0

        if not '.' in file_path:
            file_path += '.attachment'

        path = os.path.dirname(file_path) + '/'
        file_name = os.path.basename(file_path)

        if ':' in file_name:
            file_name = file_name.replace(':', '')

        while _loop:
            if os.path.isfile(file_path):
                file_path = path + str(index) + ' - ' + file_name
                index += 1
            else:
                _loop = False

        return path + file_name

    def save(self, dest_folder, backup_file_name, extension_to_dump=None):

        f = None

        if os.path.isdir(dest_folder):

            # Use long filename as first preference
            filename = self.longFilename
            if filename:
                filename = msgdump.bytes_to_windows_string(self, filename)
            # Otherwise use the short filename

            if filename is None:
                filename = self.shortFilename
            # Otherwise just make something up!
            if filename is None:
                filename = backup_file_name

            if extension_to_dump:
                if extension_to_dump in filename:
                    f = open(dest_folder + filename, 'wb')
                    f.write(self.data)
                    f.close()
            else:
                try:
                    if isinstance(filename, bytes):
                        org_filename = filename
                        filename = filename.decode('utf8', 'ignore')

                    file_path = dest_folder + filename
                    file_path = self.get_unique_filename(file_path)

                    f = open(file_path, 'wb')
                except ValueError:
                    filename = org_filename.decode('utf-16-le', 'ignore')
                    file_path = dest_folder + filename
                    file_path = self.get_unique_filename(file_path)
                    f = open(file_path, 'wb')

                try:
                    f.write(self.data)
                    f.close()
                except TypeError:
                    print('WARNING: Unable to save the attachment: %s' % file_path)

        else:
            print('Destination folder: %s -> Not found !' % dest_folder)
            return ''

        return filename

class msgdump(OleFile.OleFileIO):

    email_part = {
        'body': '__substg1.0_1000',
        'subject': '__substg1.0_0037',
        'sender': '__substg1.0_0C1F',
        'to': '__substg1.0_0E04'
    }

    def __init__(self, filename):

        self.filename = filename
        self.initialized = None

        if os.path.isdir(filename):
            self.initialized = None
            print('Is not an OLE file: %s' % filename)

        if os.path.isfile(filename):
            if not OleFile.isOleFile(filename):
                self.initialized = None
                print('Is not an OLE file or is not accessible: %s' % filename)

            OleFile.OleFileIO.__init__(self, filename)
            self.ole = OleFile.OleFileIO(filename)
            self.initialized = True

        prefix = ''
        prefixl = []
        tmp_condition = prefix != ''
        if tmp_condition:
            if not isinstance(prefix, str):
                try:
                    prefix = '/'.join(prefix)
                except:
                    raise TypeError('Invalid prefix type: ' + str(type(prefix)) +
                                    '\n(This was probably caused by you setting it manually).')
            prefix = prefix.replace('\\', '/')
            g = prefix.split("/")
            if g[-1] == '':
                g.pop()
            prefixl = g
            if prefix[-1] != '/':
                prefix += '/'
        self.__prefix = prefix
        self.__prefixList = prefixl

    def list_streams(self):
        print(*self.ole.listdir(), sep='\n')

    def bytes_to_windows_string(self, string):
        if string is None:
            return None
        return str(string, 'utf_16_le')

    def _getStream(self, name, type=''):

        if type:
            stream_name = name + type
        else:
            stream_name = name

        try:
            stream = self.openstream(stream_name)
        except OSError:
            return None

        if stream:
            return stream.read()
        else:
            return None

    def _getAttachments(self):

        attachmentDirs = []

        for dir_ in self.ole.listdir():
            if dir_[0].startswith('__attach') and dir_[0] not in attachmentDirs:
                attachmentDirs.append(dir_[0])

        self._attachments = []

        for attachmentDir in attachmentDirs:
            self._attachments.append(Attachment(self, attachmentDir))

        return self._attachments

    def _getStringStream(self, stream_name, prefer='unicode'):

        _stream_name = self.email_part.get(stream_name, stream_name)

        try:
            ascii_string = self._getStream(_stream_name, '001E') # 001E
        except OSError:
            ascii_string = None
            pass  # The ascii version not found

        try:
            unicode_string = self.bytes_to_windows_string(self._getStream(_stream_name, '001F'))  # 001F
        except OSError:
            unicode_string = None
            pass  # The ascii version not found

        if ascii_string is None:
            return unicode_string
        elif unicode_string is None:

            if isinstance(ascii_string, bytes):
                ascii_string = ascii_string.decode(errors='ignore')
            return ascii_string
        else:
            if prefer == 'unicode':
                return unicode_string
            else:
                return ascii_string

    def header(self):
        try:
            return self._header
        except Exception:
            headerText = self._getStringStream('__substg1.0_007D')
            self._header = headerText

    def recipients(self):
        """
        Returns a list of all recipients.
        """
        try:
            return self._recipients
        except AttributeError:
            # Get the recipients
            recipientDirs = []

            for dir_ in self.ole.listdir():
                if dir_[len(self.__prefixList)].startswith('__recip') and \
                                dir_[len(self.__prefixList)] not in recipientDirs:
                    recipientDirs.append(dir_[len(self.__prefixList)])

            self._recipients = []

            for recipientDir in recipientDirs:
                self._recipients.append(Recipient(recipientDir, self, self.__prefix))

            return self._recipients


class Recipient(object):
    """
    Contains the data of one of the recipients in an msg file.
    """
    def fix_path(self, inp, prefix=True):
        """
        Changes paths so that they have the proper
        prefix (should :param prefix: be True) and
        are strings rather than lists or tuples.
        """
        if isinstance(inp, (list, tuple)):
            inp = '/'.join(inp)
        if prefix:
            inp = self.prefix_value + inp
        return inp

    def windowsUnicode(self, string):
        return str(string, 'utf_16_le') if string is not None else None
        return ""

    def _getStreamA(self, filename):

        try:
            with self.__msg.openstream(filename) as stream:
                return stream.read()
        except OSError:
            # print('Stream "{}" was requested but could not be found. Returning `None`.'.format(filename))
            return None

    def _getStringStreamA(self, filename, prefer='unicode', prefix=True):
        """
        Gets a string representation of the requested filename.
        Checks for both ASCII and Unicode representations and returns
        a value if possible.  If there are both ASCII and Unicode
        versions, then :param prefer: specifies which will be
        returned.
        """

        # '__recip_version1.0_#00000000/__substg1.0_39FE001F'

        filename = self.fix_path(filename, prefix)

        asciiVersion = self._getStreamA(filename + '001E')
        unicodeVersion = self._getStreamA(filename + '001F')

        unicodeVersion = self.windowsUnicode(string=unicodeVersion)

        if asciiVersion is None:
            return unicodeVersion
        elif unicodeVersion is None:
            return asciiVersion
        else:
            if prefer == 'unicode':
                return unicodeVersion
            else:
                return asciiVersion

    def __init__(self, _dir, msg, prefix_value=''):
        object.__init__(self)
        self.prefix_value = prefix_value
        self.__msg = msg  # Allows calls to original msg file
        self.__dir = _dir
        # self.__props = Properties(self._getStream('__properties_version1.0'), TYPE_RECIPIENT)
        self.__email = self._getStringStream('__substg1.0_39FE')
        if not self.__email:
            self.__email = self._getStringStream('__substg1.0_3003')
        self.__name = self._getStringStream('__substg1.0_3001')
        # self.__type = self.__props.get('0C150003').value
        self.__formatted = u'{0} <{1}>'.format(self.__name, self.__email)

    def _getStream(self, filename):
        return self.__msg._getStream([self.__dir, filename])

    def _getStringStream(self, filename):
        """
        Gets a string representation of the requested filename.
        Checks for both ASCII and Unicode representations and returns
        a value if possible.  If there are both ASCII and Unicode
        versions, then :param prefer: specifies which will be
        returned.
        """
        return self._getStringStreamA([self.__dir, filename])

    def email(self):

        if isinstance(self.__email, bytes):
            self.__email = self.__email.decode(errors='ignore')
        return self.__email



class body_parser(object):

    def __init__(self, email_body):
        self.body = email_body

    def get_urls(self):
        urls = []
        if isinstance(self.body, str):
            urls = list(iocextract.extract_urls(self.body, refang=True))
        else:
            print('E-mail body is not a string !')

        return urls


class text_parser(object):

    category_mappings = {
        'Computer/Information Security': '108', 'For Kids': '87', 'Alcohol': '23',
        'Entertainment': '20', 'Travel': '66',
        'Proxy Avoidance': '86', 'Potentially Unwanted Software': '102',
        'Charitable Organizations': '29', 'Weapons': '15',
        'Religion': '54', 'Health': '37', 'Sexual Expression': '93',
        'File Storage/Sharing': '56', 'Gambling': '11',
        'Software Downloads': '71', 'Email': '52', 'News/Media': '46',
        'Personals/Dating': '47', 'Adult/Mature Content': '1',
        'Newsgroups/Forums': '53', 'Piracy/Copyright Concerns': '118',
        'Mixed Content/Potentially Adult': '50', 'Shopping': '58',
        'Remote Access Tools': '57', 'Business/Economy': '21', 'Informational': '107',
        'Non-Viewable/Infrastructure': '96',
        'Society/Daily Living': '61', 'Peer-to-Peer (P2P)': '83', 'Media Sharing': '112',
        'Scam/Questionable/Illegal': '9',
        'Audio/Video Clips': '84', 'Humor/Jokes': '68', 'Spam': '101',
        'Office/Business Applications': '85',
        'Political/Social Advocacy': '36', 'Internet Connected Devices': '109',
        'Translation': '95',
        'Alternative Spirituality/Belief': '22', 'Extreme': '7', 'Online Meetings': '111',
        'Sex Education': '4',
        'Web Ads/Analytics': '88', 'Technology/Internet': '38', 'Tobacco': '24',
        'Art/Culture': '30', 'Phishing': '18',
        'Intimate Apparel/Swimsuit': '5', 'Vehicles': '67', 'Abortion': '16',
        'Web Hosting': '89', 'TV/Video Streams': '114',
        'Controlled Substances': '25', 'Malicious Outbound Data/Botnets': '44', 'Games': '33',
        'Auctions': '59',
        'Brokerage/Trading': '32', 'Military': '35', 'Hacking': '17',
        'E-Card/Invitations': '106', 'Social Networking': '55',
        'Chat (IM)/SMS': '51', 'Sports/Recreation': '65', 'Search Engines/Portals': '40',
        "I Don't Know": '90', 'Job Search/Careers': '45',
        'Reference': '49', 'Content Servers': '97', 'Nudity': '6',
        'Restaurants/Dining/Food': '64', 'Suspicious': '92',
        'Child Pornography': '26', 'Marijuana': '121', 'Placeholders': '98',
        'Radio/Audio Streams': '113', 'Government/Legal': '34',
        'Financial Services': '31', 'Malicious Sources/Malnets': '43', 'Real Estate': '60',
        'Pornography': '3', 'Dynamic DNS Host': '103',
        'Education': '27', 'Internet Telephony': '110', 'Personal Sites': '63',
        'Violence/Hate/Racism': '14'
    }

    columns_to_print = ['file_path', 'type', 'tracking_id', 'subject', 'Submission_Date', 'Submitter', 'files_submitted']

    string_type_mapping = {
        '[CLOSED]: Symantec Security Response Automation': 'Symantec Submission Closure',
        'Symantec Security Response Scribe Automation': 'Symantec Scribe Report',
        'Blue Coat Site Review submission': 'Blue Coat Site Review submission'

    }

    def obfuscate(self, url):
        return iocextract.defang(url)

    def deobfuscate(self, url):
        return iocextract.refang_url(url)

    def split(self, strng, sep, pos):
        strng = strng.split(sep)
        return sep.join(strng[:pos]), sep.join(strng[pos:])

    def _get_type(self, input_string):

        for _key, _type in self.string_type_mapping.items():
            if _key in input_string:
                return _type

        return 'Unknown'

    def _get_tracking_id(self, input_string, input_type):

        if input_type == 'Symantec Submission Closure':
            _, __, tracking_id = input_string.rpartition('#')
            return tracking_id[:-1]
        elif input_type == 'Symantec Scribe Report':
            _, __, tracking_id = input_string.rpartition('#')
            return tracking_id[:-8]

        elif input_type == 'Blue Coat Site Review submission':
            _, __, tracking_id = input_string.rpartition('#')
            return tracking_id
        else:
            return 'Unknown input string'

    def _get_submission_date(self, input_string, input_type):

        if input_type in ['Symantec Submission Closure','Symantec Scribe Report']:
            _submission_date = re.findall(r'^Submission Date(.+)Tracking', input_string, re.IGNORECASE + re.MULTILINE)

            if not _submission_date:
                _submission_date = re.findall(r'([0-9]{4}-[0-9]{2}-[0-9]{2} [0-9]{2}:[0-9]{2}:[0-9]{2})', input_string, re.IGNORECASE + re.MULTILINE)
        elif input_type in 'Blue Coat Site Review submission':
            _submission_date = re.findall(r'Reviewed:.*([A-Za-z]{3,}.+?UTC)', input_string, re.IGNORECASE + re.MULTILINE)
        else:
            return 'Unknown input string'

        if _submission_date == []:
            _submission_date = 'Unknown Submission Date'
        else:
            _submission_date = _submission_date[0].strip().replace(',', ';')

        return _submission_date

    def _get_submitter(self, input_string, input_type):

        if input_type in ['Symantec Submission Closure','Symantec Scribe Report']:
            output = re.findall(r'^Submitter(.+)$', input_string, re.IGNORECASE + re.MULTILINE)

            if output == []:
                output = 'Unknown Submission Date'
            else:
                output = output[0].strip()

            return output
        else:
            return 'Unknown input string'

    def _get_files_submitted(self, input_string, input_type):

        input_string_lines = input_string.split('\n')
        start_parsing = False
        files = []
        index = 0
        start_index = None
        end_index = None

        # Determine the area to search in
        for line in input_string_lines:

            if "Files Submitted" in line:
                start_index = index
                continue

            if "Developer Notes" in line:
                end_index = index
                continue

                if start_index:
                    break

            index += 1

        _file = []
        next_line = ''

        if start_index and end_index:

            #  Hash submission
            if 'Submission Hash' in input_string:

                files_submitted_section = input_string_lines[start_index + 1:end_index - 1]
                test = re.split(r'\r|\t', "".join(files_submitted_section))
                test = [item.strip() for item in test]
                [test.remove(item) if item == '' or item == ' ' else item.strip() for item in test]

                for _line in test:
                    _line = _line.strip()

                    if 'WinUpdateexe.exe' in _line:
                        test = ""

                    if re.search('^[0-9]{1,2}$', _line, re.IGNORECASE):

                        _file.append('separator')
                        _file.append(str(_line))
                        next_line = 'filename'
                        continue

                    if next_line == 'filename':
                        _file.append(str(_line))
                        next_line = 'md5'
                        continue

                    if next_line == 'md5':
                        _file.append(str(_line))
                        next_line = 'determination'
                        continue

                    if next_line == 'determination':
                        _file.append(str(_line))
                        next_line = 'signature'
                        continue

                    if next_line == 'signature':
                        _file.append(str(_line))
                        next_line = 'RR Seq'
                        continue

                    if next_line == 'RR Seq':
                        _file.append(str(_line))
                        files.append(_file[1:])
                        _file = []
                        next_line = None

            else:
                files_submitted_section = input_string_lines[start_index + 3:end_index - 1]

                test = re.split(r'\r|\t', "".join(files_submitted_section))
                test = [item.strip() for item in test]
                [test.remove(item) if item == '' or item == ' ' else item.strip() for item in test]

                """
                files_submitted_section = "".join(files_submitted_section)
                files_submitted_section = files_submitted_section.replace('\r', ' ')
                files_submitted_section = re.sub(r'\t{1,}', '\t', files_submitted_section)
                files_submitted_section = re.sub(r'\t {1}', '\t', files_submitted_section)
                files_submitted_section = re.sub(r' {2}', '\t', files_submitted_section)
                files_submitted_section = re.sub(r'\t ', '\t', files_submitted_section)
                files_submitted_section = re.sub(r' \t', '\t', files_submitted_section)
                files_submitted_section = files_submitted_section.split('\t')
                """

                index = 0
                status_ok = False
                while not status_ok:
                    #if index > len(files_submitted_section) or index > 100:
                    if index > len(test) or index > 100:
                        status_ok = True

                    try:
                        #_line = files_submitted_section[index].strip()
                        _line = test[index].strip()

                        if '3' in _line:
                            X = ""

                    except IndexError:
                        status_ok = True
                        break

                    if re.search('^[0-9]{1,2}$', _line, re.IGNORECASE + re.MULTILINE):

                        #  Check of the record is corrupted (contains next file)
                        #_items = files_submitted_section[index:index + 6]
                        _items = test[index:index + 7]
                        _file.extend(_items)

                        try:
                            last_item = _file[-1:][0].strip()
                            look_behind = _file[-2:][0].strip()
                        except IndexError:
                            X = ""

                        pos_of_next_file = re.search('^[0-9]{1,2}$', last_item, re.IGNORECASE)
                        if not pos_of_next_file:
                            pos_of_next_file = re.search('^[0-9]{1,2} ', look_behind, re.IGNORECASE)

                        # FIX:  Due to inconsistent body content. The hash lands in the index
                        if ' ' in _file[0]:
                            _index, __, file_name = _file[0].partition(' ')

                            _file[0] = _index
                            _file[1] = file_name

                        #  FIX:  Due to inconsistent body content. The hash lands in the file name
                        if ' ' in _file[1]:
                            file_name, __, _hash = _file[1].rpartition(' ')

                            if len(_hash) == 32:
                                _file[1] = file_name
                                _file.insert(2, _hash)

                        #  FIX1:  Due to inconsistent body content, a tab between hash and determination is a space!
                        if ' ' in _file[2]:
                            fields = _file[2].split(' ', maxsplit=1)
                            _file[2] = fields[0]
                            _file.insert(3, fields[1])


                        if not ' ' in _file[3] and any(determination in _file[3] for determination in ['Not', 'Data',
                                                                               'Threat', 'New',
                                                                               'Already']):
                            _file[3] = _file[3] + ' ' + _file[4]
                            _file.remove(_file[4])


                        # FIX2:  Due to inconsistent body content:
                        #  Determination column, contains unexpected space instead of a tab
                        if ' ' in _file[3]:
                            correct_format = None
                            if any(determination in _file[3] for determination in ['Not Malicious', 'Data File',
                                                                                   'Threat artifact', 'New Threat',
                                                                                   'Already Detected']):
                                if _file[3].count(' ') > 1:
                                    fields = self.split(_file[3], ' ', 2)
                                    correct_format = False
                                else:
                                    correct_format = True
                            else:
                                fields = _file[3].split(' ', maxsplit=1)

                            if not correct_format:
                                _file[3] = fields[0]
                                _file.insert(4, fields[1])

                        # FIX3:  Due to inconsistent body content, If a detection name is N/A, Symantec uses a
                        #  space instead of tab between file name and a detection name
                        if ' N/A' in _file[3]:
                            _file[3] = _file[3].replace(' N/A', '')
                            _file.append('N/A')

                        if pos_of_next_file:
                            index = index + 5
                            _file = _file[:-1]
                            #_file = _file[0:5]
                            files.append(_file)
                        else:
                            index += 6
                            files.append(_file[0:6])

                        _file = []

                    else:
                        index += 1

            return files

        else:
            return []

    def map_catid_to_categorization(self, catid):

        for key, value in self.category_mappings.items():

            if value == catid:
                return key

        return 'Unknown category id: %s' % catid

    def _get_proxy_requested_categoization(self, input_string, input_type):

        if input_type in 'Blue Coat Site Review submission':

            input_string = input_string.replace('\n', '')
            input_string = input_string.replace('\r', ' ')

            _submission_date = re.findall(r'Suggested categories:(.*?)Your', input_string, re.IGNORECASE + re.MULTILINE)

            if not _submission_date:
                _submission_date = re.findall(r'Suggested category:(.*?)Your',
                                              input_string, re.IGNORECASE + re.MULTILINE)

        else:
            return 'Unknown input string'

        if _submission_date == []:
            _submission_date = 'Unknown Requested Proxy Categorization'
        else:

            cat_nums = re.findall(r'catnum=(\d+)\&', _submission_date[0], re.IGNORECASE + re.MULTILINE)
            _submission_date = []

            for cat_id in cat_nums:
                _submission_date.append(self.map_catid_to_categorization(cat_id))

            _submission_date = '[%s]' % ';'.join(_submission_date)

        return _submission_date

    def _get_proxy_categoization(self, input_string, input_type):

        if input_type in 'Blue Coat Site Review submission':

            input_string = input_string.replace('\n', '')
            input_string = input_string.replace('\r', ' ')

            _submission_date = re.findall(r'URL as(.*?)\. ', input_string, re.IGNORECASE + re.MULTILINE)

            if not _submission_date:
                _submission_date = re.findall(r' as(.*?)\. ', input_string, re.IGNORECASE + re.MULTILINE)

        else:
            return 'Unknown input string'

        if _submission_date == []:
            _submission_date = 'Unknown Proxy Categorization'
        else:

            cat_nums = re.findall(r'catnum=(\d+)\&', _submission_date[0], re.IGNORECASE + re.MULTILINE)
            _submission_date = []

            for cat_id in cat_nums:
                _submission_date.append(self.map_catid_to_categorization(cat_id))

            _submission_date = '[%s]' % ';'.join(_submission_date)

        return _submission_date

    def _get_submitted_url(self, input_string, input_type):

        if input_type in 'Blue Coat Site Review submission':

            input_string = input_string.replace('\n', '')
            input_string = input_string.replace('\r', ' ')

            _submission_date = re.findall(r'Submitted URL:(.+)Suggested', input_string, re.IGNORECASE + re.MULTILINE)
        else:
            return 'Unknown input string'

        if _submission_date == []:
            _submission_date = 'Unable to find URL'
        else:

            _submission_date = '%s' % ' '.join(_submission_date).strip()

        return self.obfuscate(url=_submission_date)

    def __init__(self, file_path, mail_subject, mail_body, mail_attachments):

        self.result = []
        row = {}

        #  Add file path
        _file_path = file_path
        row['file_path'] = _file_path

        #  Determine the e-mail type
        _type = self._get_type(input_string=mail_subject)
        row['type'] = _type

        #  Get tracking id
        _tracking_id = self._get_tracking_id(input_string=mail_subject, input_type=_type)
        row['tracking_id'] = _tracking_id

        #  E-mail subject
        _subject = mail_subject
        row['subject'] = _subject

        #  Get submission date
        _submission_date = self._get_submission_date(input_string=mail_body, input_type=_type)
        row['Submission_Date'] = _submission_date

        #  Get submitter
        _submitter = self._get_submitter(input_string=mail_body, input_type=_type)
        row['Submitter'] = _submitter

        _files_submitted = self._get_files_submitted(input_string=mail_body, input_type=_type)
        row['files_submitted'] = _files_submitted

        _proxy_categorization = self._get_proxy_categoization(input_string=mail_body, input_type=_type)

        row['proxy_category'] = _proxy_categorization

        _proxy_requested_categorization = self._get_proxy_requested_categoization(input_string=mail_body, input_type=_type)

        row['requested_proxy_category'] = _proxy_requested_categorization

        submitted_url = self._get_submitted_url(input_string=mail_body, input_type=_type)

        row['submitted_url'] = submitted_url

        self.result.append(row)



def print_raw_items(items):

    if items:
        print(*items, sep='\n')

def print_submitted_files(items):

    if items:
        for item in items:
            submitted_files = item.get('files_submitted', [])
            for files in submitted_files:
                submission_date = item.get('Submission_Date', '')
                files.append(submission_date)
                tracking_id = item.get('tracking_id', '')
                files.insert(0, tracking_id)

                # Strip the elements of the list
                map(str.strip, files)

                print(*files, sep=', ')

def print_csv(items, columns):

    row = []

    for item in items:
        for col in columns:
            row.append(item.get(col, ''))

        print(','.join(row))

        row = []

def _save_attachments(dest_folder, attachments, backup_file_name, extension_to_dump=None):

    for attachment in attachments:
        attachment.save(dest_folder=dest_folder, backup_file_name=backup_file_name, extension_to_dump=extension_to_dump)


def get_input_files(input_path, file_extensions=None, recursive=True):

    if file_extensions is None:
        file_extensions = []
    else:
        file_extensions = list(map(str.lower, file_extensions))

    input_files = []

    if input_path:

        #  Case: The input_path is an existing file
        if isfile(input_path):
            file_name, _, file_extension = input_path.rpartition('.')
            file_extension = '.' + file_extension.upper()
            input_files.append(input_path)

        #  Case: The input_path is a folder to lookup
        elif isdir(input_path):

            # Make sure tha the folder path is finished with / or \, hence the glob ** would work properly
            if 'Darwin' in _os.platform() or 'Linux' in _os.platform():
                if not input_path[:-1] == r'/':
                    input_path += r'/'

            elif 'Windows' in _os.platform():
                if not input_path[:-1] == '\\':
                    input_path += '\\'

            if recursive:
                pattern = input_path + '**'
            else:
                pattern = input_path + '*'

            for _path in glob.glob(pattern, recursive=recursive):

                #  Add only supported hives
                if isfile(_path):
                    file_name, _, file_extension = _path.rpartition('.')
                    file_extension = '.' + file_extension.upper()

                    if file_extension.lower() in file_extensions:
                        input_files.append(_path)

        if input_files:
            input_files = set(input_files)
            input_files = list(input_files)

    return input_files

def main(argv):

    # -i samples/scribe/*.msg --dump-folder dump/ --dump-attachments --dump-extension pdf -h
    argsparser = argparse.ArgumentParser(usage=argparse.SUPPRESS, description='MSGDump - Dumps info from .msg files')

    """ Argument groups """
    script_args = argsparser.add_argument_group('Script arguments', "\n")

    """ Script arguments """
    script_args.add_argument("-i", "--input", type=str, action='store', dest='input', required=True,
                             help="Input file or folder containing .msg files")

    script_args.add_argument("-o", "--output", type=str, action='store', dest='output_file', required=False,
                             help="Output files instead of stdout")

    script_args.add_argument("--recursive", action='store_true', dest='recursive_search', required=False,
                             default=False, help="Enables recursive search")

    script_args.add_argument("--raw", action='store_true', dest='print_raw_items', required=False,
                             default=False, help="Print parsed data in raw format")

    script_args.add_argument("--symc-submissions", action='store_true', dest='symc_print_submissions', required=False,
                             default=False, help="Parse submission closure e-mails from Symantec Security Response")

    script_args.add_argument("--proxy-submissions", action='store_true', dest='proxy_print_submissions', required=False,
                             default=False, help="Parse Blue Coat Site Review submission e-mails")

    script_args.add_argument("-da", "--dump-attachments", action='store_true', dest='dump_attachments', required=False,
                             default=False, help="Dump attachments")

    script_args.add_argument("-du", "--dump-urls", action='store_true', dest='dump_urls', required=False,
                             default=False, help="Dump URLs from e-mail body")

    script_args.add_argument("-df", "--dump-folder", action='store', dest='dump_folder', required=False, default=False,
                             help="Folder where dumped attachments would be saved")

    script_args.add_argument("-de", "--dump-extension", action='store', dest='dump_extension',
                             required=False, default=None, help="Dump only given extension")


    args = argsparser.parse_args()
    argc = argv.__len__()


    rows = []
    write_csv_header = True

    for filename in get_input_files(input_path=args.input, file_extensions=['.msg'], recursive=args.recursive_search):

        mail_parser = msgdump(filename=filename)
        _filename = os.path.basename(filename)
        if mail_parser.initialized is None:
            continue

        subject = mail_parser._getStringStream('subject')
        sender = mail_parser._getStringStream('sender')
        to = mail_parser._getStringStream('to')

        body = mail_parser._getStringStream('body')

        attachments = mail_parser._getAttachments()

        recipients = []
        recipients_obj = mail_parser.recipients()

        for recipient_obj in recipients_obj:
            recipients.append(recipient_obj.email())

        recipients = ';'.join(recipients)

        if not attachments:
            attachments = []

        if args.dump_urls:
            if body:

                if write_csv_header:
                    rows.append('filename,sender,recipients,subject,urls')
                    write_csv_header = False

                entry = []

                entry.append(_filename)
                entry.append(sender)
                entry.append(recipients)
                entry.append(str(subject))

                _bparser = body_parser(body)
                urls = _bparser.get_urls()
                entry.append(';'.join(urls))

                try:
                    rows.append(','.join(entry))
                except Exception:
                    debug = ""
            else:
                print('No body to parse...')

        elif args.symc_print_submissions:
            output = text_parser(file_path=filename, mail_subject=subject, mail_body=body, mail_attachments=attachments)
            rows.extend(output.result)

        elif args.proxy_print_submissions:
            output = None
            output = text_parser(file_path=filename, mail_subject=subject, mail_body=body, mail_attachments=attachments)
            rows.extend(output.result)

        if args.dump_attachments:
            _save_attachments(dest_folder=args.dump_folder, attachments=attachments, backup_file_name='XXX.sample',
                              extension_to_dump=args.dump_extension)

        if args.print_raw_items:
            print_raw_items(items=rows)

        if args.output_file:

            with open(args.output_file, "a", errors='ignore') as ofile:

                for row in rows:
                    ofile.write(row + '\n')

        elif args.proxy_print_submissions:

            if write_csv_header:
                print('file_path,tracking_id,Reviewed,requested_proxy_category,proxy_category,submitted_url')
                write_csv_header = False

            columns = ['file_path', 'tracking_id', 'Submission_Date', 'requested_proxy_category', 'proxy_category', 'submitted_url']
            print_csv(output.result, columns)

        if args.symc_print_submissions:
            print_submitted_files(items=rows)

        rows = []



if __name__ == "__main__":
    main(sys.argv)

