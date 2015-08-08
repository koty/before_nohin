# -*- coding: utf-8 -*-

__author__ = 'koty'
from zipfile import *
import chardet

"""
zip内のファイル名の文字化け対策
http://d.hatena.ne.jp/yatt/20110326/1301123145
"""

def cp932_invert(cp932_path):
    from string import printable
    def fun(i):
        pred = i > 0 and cp932_path[i] == '/' and cp932_path[i-1] not in printable
        return '\\' if pred else cp932_path[i]
    lst = map(fun, range(len(cp932_path)))
    uni = ''.join(lst).decode('cp932').replace('\\', '/')
    return uni

def as_unicode_path(path):
    #if type(path) is unicode:
    #    return path

    try:
        path_bytes = path.encode('cp437')
        detect_enc = chardet.detect(path_bytes)
        if detect_enc['confidence'] > 0:
            return path_bytes.decode(detect_enc['encoding'])
        else:
            return path_bytes.decode('SHIFT_JIS')
    except:
        pass

    # assume cp932 encoding including dame-moji
    try:
        return cp932_invert(path)
    except:
        return path

def __setattr__(self, name, value):
    if name == 'filename':
        value = as_unicode_path(value)
    object.__setattr__(self, name, value)

ZipInfo.__setattr__ = __setattr__

def infolist(self):
    return filter(lambda info: info.filename[-1] != '/', self.filelist)
ZipFile.infolist = infolist

def namelist(self):
    return [info.filename for info in self.infolist()]
    # これだとうまく動かない。何ででしょう。
    # return map(lambda info: info.filename, self.infolist())
ZipFile.namelist = namelist
