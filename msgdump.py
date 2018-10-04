__author__  = "Witold Lawacz (wit0k)"
__date__    = "2018-10-04"
__version__ = '0.0.4'

import olefile as OleFile  # pip install olefile
import glob
import re
import argparse
import sys
import os.path

""" Set working directory so the script can be executed from any location/symlink """
os.chdir(os.path.dirname(os.path.abspath(__file__)))

CRED = '\033[91m'
CYELLOW = '\33[33m'
CEND = '\033[0m'

class Attachment:

    def __init__(self, msg, dir_):
        # Get long filename
        self.longFilename = msg._getStream([dir_, '__substg1.0_3707001F'])

        # Get short filename
        self.shortFilename = msg._getStream([dir_, '__substg1.0_3704001F'])

        # Get attachment data
        self.data = msg._getStream([dir_,'__substg1.0_37010102'])

    def save(self, backup_file_name):
        # Use long filename as first preference
        filename = self.longFilename
        # Otherwise use the short filename
        if filename is None:
            filename = self.shortFilename
        # Otherwise just make something up!
        if filename is None:
            filename = backup_file_name

        f = open(filename, 'wb')
        f.write(self.data)
        f.close()
        return filename

class msgdump(OleFile.OleFileIO):

    email_part = {
        'body': '__substg1.0_1000',
        'subject': '__substg1.0_0037'
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
            return ascii_string
        else:
            if prefer == 'unicode':
                return unicode_string
            else:
                return ascii_string

class text_parser(object):

    columns_to_print = ['file_path', 'type', 'tracking_id', 'subject', 'Submission_Date', 'Submitter', 'files_submitted']

    string_type_mapping = {
        '[CLOSED]: Symantec Security Response Automation': 'Symantec Submission Closure'
    }

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
        else:
            return 'Unknown input string'

    def _get_submission_date(self, input_string, input_type):

        if input_type == 'Symantec Submission Closure':
            _submission_date = re.findall(r'^Submission Date(.+)$', input_string, re.IGNORECASE + re.MULTILINE)

            if _submission_date == []:
                _submission_date = 'Unknown Submission Date'
            else:
                _submission_date = _submission_date[0].strip()

            return _submission_date
        else:
            return 'Unknown input string'

    def _get_submitter(self, input_string, input_type):

        if input_type == 'Symantec Submission Closure':
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

                for _line in files_submitted_section:
                    _line = _line.strip()

                    if re.search('^[0-9]{1,2}$', _line, re.IGNORECASE):

                        _file.append('separator')
                        _file.append(str(_line))
                        next_line = 'filename'
                        continue

                    if next_line == 'filename':
                        _file.append(_line)
                        next_line = 'parse_details'
                        continue

                    if next_line == 'parse_details':
                        items = _line.split('\t')
                        for _item in items:
                            _file.append(_item.strip())

                        next_line = 'RR Seq'
                        continue

                    if next_line == 'RR Seq':
                        _file.append(_line)
                        files.append(_file[1:])
                        _file = []

            else:
                files_submitted_section = input_string_lines[start_index + 3:end_index - 1]
                files_submitted_section = "".join(files_submitted_section)
                files_submitted_section = files_submitted_section.replace('\r', ' ')
                files_submitted_section = re.sub(r'\t{1,}', '\t', files_submitted_section)
                files_submitted_section = re.sub(r'\t {1}', '\t', files_submitted_section)
                files_submitted_section = re.sub(r' {2}', '\t', files_submitted_section)
                files_submitted_section = re.sub(r'\t ', '\t', files_submitted_section)
                files_submitted_section = re.sub(r' \t', '\t', files_submitted_section)
                files_submitted_section = files_submitted_section.split('\t')

                index = 0
                status_ok = False
                while not status_ok:
                    if index > len(files_submitted_section) or index > 100:
                        status_ok = True

                    try:
                        _line = files_submitted_section[index].strip()
                    except IndexError:
                        status_ok = True
                        break

                    if re.search('^[0-9]{1,2}', _line, re.IGNORECASE + re.MULTILINE):

                        #  Check of the record is corrupted (contains next file)
                        _items = files_submitted_section[index:index + 6]
                        _file.extend(_items)

                        try:
                            last_item = _file[-1:][0].strip()
                        except IndexError:
                            test = ""

                        pos_of_next_file = re.search('^[0-9]{1,2}$', last_item, re.IGNORECASE)

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
                        else:
                            index += 6

                        files.append(_file[0:6])
                        _file = []

                    else:
                        index += 1

            return files

        else:
            return []



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

        self.result.append(row)

def print_raw_items(items):

    if items:
        print(CYELLOW + 'Raw Items:' + CEND)
        print(*items, sep='\n')

def print_submitted_files(items):

    if items:
        print(CYELLOW + 'Symantec Submissions:' + CEND)
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

def main(argv):

    argsparser = argparse.ArgumentParser(usage=argparse.SUPPRESS, description='MSGDump - Dumps info from .msg files')

    """ Argument groups """
    script_args = argsparser.add_argument_group('Script arguments', "\n")

    """ Script arguments """
    script_args.add_argument("-i", "--input", type=str, action='store', dest='input', required=True,
                             help="Input file or folder pattern like samples/*.msg")

    script_args.add_argument("--raw", action='store_true', dest='print_raw_items', required=False,
                             default=False, help="Print parsed data in raw format")

    script_args.add_argument("--symc-submissions", action='store_true', dest='symc_print_submissions', required=False,
                             default=False, help="Print details according to file submissions sent to Symantec")

    args = argsparser.parse_args()
    argc = argv.__len__()


    rows = []

    for filename in glob.glob(args.input):

        mail_parser = msgdump(filename=filename)

        if mail_parser.initialized is None:
            continue

        subject = mail_parser._getStringStream('subject')
        body = mail_parser._getStringStream('body')
        attachments = mail_parser._getAttachments()

        if not attachments:
            attachments = []

        output = text_parser(file_path=filename, mail_subject=subject, mail_body=body, mail_attachments=attachments)
        rows.extend(output.result)

    if args.print_raw_items:
        print_raw_items(items=rows)

    if args.symc_print_submissions:
        print_submitted_files(items=rows)

if __name__ == "__main__":
    main(sys.argv)

