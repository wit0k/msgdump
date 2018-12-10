__author__  = "Witold Lawacz (wit0k)"
__date__    = "2018-10-04"
__version__ = '0.1.0.1'

import olefile as OleFile  # pip install olefile
import glob
import re
import argparse
import sys
import os.path
import iocextract

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

    columns_to_print = ['file_path', 'type', 'tracking_id', 'subject', 'Submission_Date', 'Submitter', 'files_submitted']

    string_type_mapping = {
        '[CLOSED]: Symantec Security Response Automation': 'Symantec Submission Closure',
        'Symantec Security Response Scribe Automation': 'Symantec Scribe Report'

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
        elif input_type == 'Symantec Scribe Report':
            _, __, tracking_id = input_string.rpartition('#')
            return tracking_id[:-8]
        else:
            return 'Unknown input string'

    def _get_submission_date(self, input_string, input_type):

        if input_type in ['Symantec Submission Closure','Symantec Scribe Report']:
            _submission_date = re.findall(r'^Submission Date(.+)Tracking', input_string, re.IGNORECASE + re.MULTILINE)

            if not _submission_date:
                _submission_date = re.findall(r'([0-9]{4}-[0-9]{2}-[0-9]{2} [0-9]{2}:[0-9]{2}:[0-9]{2})', input_string, re.IGNORECASE + re.MULTILINE)

            if _submission_date == []:
                _submission_date = 'Unknown Submission Date'
            else:
                _submission_date = _submission_date[0].strip()
        else:
            return 'Unknown input string'

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


def _save_attachments(dest_folder, attachments, backup_file_name, extension_to_dump=None):

    for attachment in attachments:
        attachment.save(dest_folder=dest_folder, backup_file_name=backup_file_name, extension_to_dump=extension_to_dump)


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

    for filename in glob.glob(args.input):

        mail_parser = msgdump(filename=filename)
        _filename = os.path.basename(filename)
        if mail_parser.initialized is None:
            continue

        subject = mail_parser._getStringStream('subject')
        body = mail_parser._getStringStream('body')
        attachments = mail_parser._getAttachments()

        if not attachments:
            attachments = []

        if args.dump_urls:
            if body:
                entry = []
                entry.append(_filename)
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

        if args.dump_attachments:
            _save_attachments(dest_folder=args.dump_folder, attachments=attachments, backup_file_name='XXX.sample',
                              extension_to_dump=args.dump_extension)

        if args.print_raw_items:
            print_raw_items(items=rows)

        if args.symc_print_submissions:
            print_submitted_files(items=rows)

        rows = []



if __name__ == "__main__":
    main(sys.argv)

