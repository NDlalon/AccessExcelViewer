import codecs
import tempfile

icon=b''

def getIcon_tempFile():
    file=tempfile.NamedTemporaryFile(delete=False, suffix='.ico')
    file.write(icon)

    return file.name

def getIcon_bin():
    return icon
