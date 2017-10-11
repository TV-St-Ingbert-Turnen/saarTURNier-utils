VERSION = '0.1'

def check_version(doc_version_string):
    doc_major, doc_minor = tuple([int(x) for x in doc_version_string[1:].split('.')])
    lib_major, lib_minor = tuple([int(x) for x in VERSION.split('.')])

    is_matching = doc_major == lib_major and doc_minor == lib_minor

    if not is_matching:
        raise IOError("Document version {} is not supported, the version of this library is v{}"
                      .format(doc_version_string, VERSION))
